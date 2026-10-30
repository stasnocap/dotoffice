using DotOffice.Converter;
using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Pdf;
using DotOffice.Samples;

// Usage:
//   DotOffice.Samples convert <input.docx> <output.pdf> [--fonts <dir>] [--no-system-fonts]
//   DotOffice.Samples showcase <output.docx> [image.jpg]
//   DotOffice.Samples <font.ttf> [image.jpg] [output.pdf]      (low-level PDF writer demo)
if (args.Length < 1)
{
    Console.Error.WriteLine("Usage: DotOffice.Samples convert <input.docx> <output.pdf> [--fonts <dir>] [--no-system-fonts]");
    Console.Error.WriteLine("       DotOffice.Samples showcase <output.docx> [image.jpg]");
    Console.Error.WriteLine("       DotOffice.Samples <font.ttf> [image.jpg] [output.pdf]");
    return 1;
}

switch (args[0])
{
    case "convert" when args.Length >= 3:
        return await Convert(args[1], args[2], args[3..]);
    case "showcase" when args.Length >= 2:
        ShowcaseDocx.Write(args[1], args.Length > 2 ? File.ReadAllBytes(args[2]) : null);
        Console.WriteLine($"Wrote {args[1]}");
        return 0;
    default:
        return HelloPdf(args);
}

static async Task<int> Convert(string input, string output, string[] flags)
{
    var warnings = new ConversionWarningList();
    FontCollection? fonts = null;
    for (int i = 0; i < flags.Length; i++)
        if (flags[i] == "--fonts" && i + 1 < flags.Length) fonts = new FontCollection().AddDirectory(flags[++i]);

    var options = new ConvertOptions
    {
        Fonts = fonts,
        UseSystemFonts = !flags.Contains("--no-system-fonts"),
        Warnings = warnings,
    };
    var sw = System.Diagnostics.Stopwatch.StartNew();
    await DocumentConverter.ConvertAsync(input, output, options);
    Console.WriteLine($"Wrote {output} ({new FileInfo(output).Length:N0} bytes) in {sw.ElapsedMilliseconds} ms");
    foreach (var warning in warnings.Warnings) Console.WriteLine($"  warning {warning}");
    return 0;
}

static int HelloPdf(string[] args)
{
    var output = args.Length > 2 ? args[2] : "hello.pdf";
    var doc = new PdfDocument { Title = "DotOffice hello world", Author = "DotOffice" };
    var font = doc.AddFont(TrueTypeFont.Load(args[0]));
    var page = doc.AddPage(PageSize.A4);
    var canvas = page.Canvas;

    const double margin = 72;
    double y = page.Size.Height - margin - 24;
    canvas.SetFillColor(0.12, 0.47, 0.71).DrawText(font, 28, margin, y, "Hello from DotOffice");

    y -= 40;
    canvas.SetFillColor(0, 0, 0);
    foreach (var line in new[] { "Pure managed .NET — no LibreOffice.", "Привет, мир! Γειά σου κόσμε!", "Accents: café, naïve, Ångström" })
    {
        canvas.DrawText(font, 14, margin, y, line);
        y -= 22;
    }

    const string measured = "This line is underlined using measured width.";
    y -= 10;
    canvas.DrawText(font, 12, margin, y, measured)
          .SetStrokeColor(1, 0.5, 0.05).SetLineWidth(1)
          .MoveTo(margin, y - 3).LineTo(margin + font.MeasureText(measured, 12), y - 3).Stroke();

    y -= 30;
    canvas.SetStrokeColor(0.5, 0.5, 0.5).Rectangle(margin, y - 60, page.Size.Width - 2 * margin, 60).Stroke();

    if (args.Length > 1)
    {
        var image = doc.AddImage(PdfImage.FromJpeg(File.ReadAllBytes(args[1])));
        canvas.DrawImage(image, margin, y - 200, image.Width, image.Height);
    }

    using (var stream = File.Create(output))
        doc.Save(stream);

    Console.WriteLine($"Wrote {output} ({new FileInfo(output).Length:N0} bytes)");
    return 0;
}