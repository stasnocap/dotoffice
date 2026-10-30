using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using DotOffice.Converter;
using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Tests.Docx;
using static DotOffice.Tests.Docx.DocxBuilder;

namespace DotOffice.Tests;

public class DocumentConverterTests
{
    private static readonly FontCollection TestFonts = new FontCollection().AddFile(TestAssets.Path("NotoSans-Regular.ttf"));

    private static async Task<byte[]> Convert(DocxBuilder docx, ConvertOptions? options = null)
    {
        using var output = new MemoryStream();
        await DocumentConverter.ConvertAsync(docx.BuildStream(), output, DocumentFormat.Docx, DocumentFormat.Pdf,
            options ?? new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false });
        return output.ToArray();
    }

    /// <summary>A document touching most features: styles, lists, a table, header/footer, images, two sections.</summary>
    private static DocxBuilder Sample()
    {
        var b = new DocxBuilder
        {
            Title = "E2E",
            Styles = "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
                     "<w:style w:type=\"paragraph\" w:styleId=\"H1\"><w:pPr><w:keepNext/><w:spacing w:before=\"240\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>",
            Numbering = "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/>" +
                        "<w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>",
        };
        var png = b.AddImage("a.png", PngTestEncoder.Encode(4, 4, 6, 8, Enumerable.Range(0, 64).Select(i => i * 4).ToArray()));
        var jpg = b.AddImage("b.jpeg", TestAssets.Bytes("gradient.jpg"));
        var header = b.AddHeader(P(R("Running header")));
        var footer = b.AddFooter(P("<w:r><w:t xml:space=\"preserve\">Page </w:t></w:r><w:fldSimple w:instr=\"PAGE\"><w:r><w:t>1</w:t></w:r></w:fldSimple>" +
                                   "<w:r><w:t xml:space=\"preserve\"> of </w:t></w:r><w:fldSimple w:instr=\"NUMPAGES\"><w:r><w:t>1</w:t></w:r></w:fldSimple>", "<w:jc w:val=\"center\"/>"));
        var body = new StringBuilder();
        body.Append(P(R("Introduction"), "<w:pStyle w:val=\"H1\"/>"));
        body.Append(P(R("First list item"), "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"));
        body.Append(P(R("Second list item"), "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"));
        body.Append("<w:tbl><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/><w:bottom w:val=\"single\" w:sz=\"4\"/><w:insideH w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPr>" +
                    "<w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>" +
                    $"<w:tr><w:tc>{P(R("Cell A1"))}</w:tc><w:tc>{P(R("Cell B1"))}</w:tc></w:tr>" +
                    $"<w:tr><w:tc>{P(R("Cell A2"))}</w:tc><w:tc>{P(R("Cell B2"))}</w:tc></w:tr></w:tbl>");
        body.Append(P(InlineImage(png, 40, 40) + R(" and ") + InlineImage(jpg, 80, 50)));
        for (int i = 1; i <= 60; i++) body.Append(P(R($"Filler paragraph {i} with enough words to wrap across the line at least once in the column."), "<w:jc w:val=\"both\"/>"));
        body.Append(P(R("Last portrait paragraph"), Section(extra: $"<w:headerReference w:type=\"default\" r:id=\"{header}\"/><w:footerReference w:type=\"default\" r:id=\"{footer}\"/>")));
        body.Append(P(R("Landscape page")));
        body.Append(Section(15840, 12240));
        b.Body = body.ToString();
        return b;
    }

    [Fact]
    public async Task Converts_docx_to_a_valid_pdf()
    {
        var pdf = await Convert(Sample());
        var text = Encoding.Latin1.GetString(pdf);
        Assert.StartsWith("%PDF-1.7", text);
        int pages = Regex.Matches(text, @"/Type /Page\b").Count;
        Assert.True(pages >= 3, $"expected several pages, got {pages}");
        Assert.Contains("/MediaBox [0 0 792 612]", text);   // landscape section
        Assert.Contains("/SMask", text);                     // PNG alpha
        Assert.Contains("/Filter /DCTDecode", text);         // JPEG passthrough
        Assert.Contains("/FontFile2", text);
        Assert.Contains("/Title (E2E)", text);

        var path = Path.Combine(Path.GetTempPath(), $"dotoffice-{Guid.NewGuid():N}.pdf");
        await File.WriteAllBytesAsync(path, pdf);
        try
        {
            if (Tool("qpdf", "--check", path) is { } check)
                Assert.True(check.ExitCode == 0, check.Output);
            if (Tool("pdftotext", "-layout", path, "-") is { } extracted)
            {
                var t = extracted.Output;
                Assert.Contains("Running header", t);
                string[] ordered = ["Introduction", "1.", "First list item", "2.", "Second list item", "Cell A1", "Cell B1", "Cell A2", "Filler paragraph 1 ", "Last portrait paragraph", "Landscape page"];
                int at = 0;
                foreach (var s in ordered)
                {
                    int found = t.IndexOf(s, at, StringComparison.Ordinal);
                    Assert.True(found >= 0, $"'{s}' missing or out of order");
                    at = found + s.Length;
                }
                Assert.Matches($@"Page 1 of {pages}\b", t);
                Assert.Matches($@"Page {pages} of {pages}\b", t);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task Reports_warnings_and_can_skip_font_embedding()
    {
        var warnings = new ConversionWarningList();
        var pdf = await Convert(Sample(), new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false, Warnings = warnings, EmbedFonts = false });
        var text = Encoding.Latin1.GetString(pdf);
        Assert.DoesNotContain("/FontFile2", text);
        Assert.Contains(warnings.Warnings, w => w.Code == "font.substituted" && w.Message.Contains("Calibri"));
    }

    [Fact]
    public async Task Unsupported_images_become_placeholders()
    {
        var b = new DocxBuilder();
        var gif = b.AddImage("x.gif", "GIF89a\x01\0\x01\0\0\0\0;"u8.ToArray());
        b.Body = P(InlineImage(gif, 30, 30));
        var warnings = new ConversionWarningList();
        var pdf = await Convert(b, new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false, Warnings = warnings });
        Assert.StartsWith("%PDF", Encoding.Latin1.GetString(pdf));
        Assert.Contains(warnings.Warnings, w => w.Code == "pdf.image-format");
    }

    [Fact]
    public async Task Unsupported_conversion_throws()
    {
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            DocumentConverter.ConvertAsync(Stream.Null, Stream.Null, DocumentFormat.Xlsx, DocumentFormat.Pdf));
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            DocumentConverter.ConvertAsync(Stream.Null, Stream.Null, DocumentFormat.Docx, DocumentFormat.Docx));
    }

    [Fact]
    public async Task Invalid_input_throws_and_leaves_no_output_file()
    {
        var input = Path.Combine(Path.GetTempPath(), $"dotoffice-{Guid.NewGuid():N}.docx");
        var output = Path.ChangeExtension(input, ".pdf");
        await File.WriteAllBytesAsync(input, [1, 2, 3]);
        try
        {
            await Assert.ThrowsAsync<InvalidDataException>(() => DocumentConverter.ConvertAsync(input, output));
            Assert.False(File.Exists(output));
        }
        finally
        {
            File.Delete(input);
        }
    }

    [Fact]
    public async Task Requires_a_font_source()
    {
        await Assert.ThrowsAsync<ArgumentException>(() => DocumentConverter.ConvertAsync(Sample().BuildStream(), Stream.Null,
            DocumentFormat.Docx, DocumentFormat.Pdf, new ConvertOptions { UseSystemFonts = false }));
    }

    [Fact]
    public async Task Honors_cancellation()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocumentConverter.ConvertAsync(Sample().BuildStream(), Stream.Null,
            DocumentFormat.Docx, DocumentFormat.Pdf, new ConvertOptions { Fonts = TestFonts, UseSystemFonts = false }, cts.Token));
    }

    /// <summary>Runs an external tool if it is installed (CI installs qpdf/poppler on Linux); null otherwise.</summary>
    private static (int ExitCode, string Output)? Tool(string name, params string[] args)
    {
        if (OperatingSystem.IsWindows()) return null;
        try
        {
            var psi = new ProcessStartInfo(name) { RedirectStandardOutput = true, RedirectStandardError = true };
            foreach (var a in args) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi)!;
            var output = p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
            p.WaitForExit();
            return (p.ExitCode, output);
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return null;
        }
    }
}