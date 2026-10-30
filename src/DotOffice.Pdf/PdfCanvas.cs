using System.Text;

namespace DotOffice.Pdf;

/// <summary>
/// Builds a page content stream. Coordinates are in points with the origin at the
/// bottom-left corner of the page (native PDF space).
/// </summary>
public sealed class PdfCanvas
{
    private readonly StringBuilder _content = new();

    internal HashSet<PdfFont> Fonts { get; } = [];
    internal HashSet<PdfImage> Images { get; } = [];

    public PdfCanvas SaveState() => Op("q");
    public PdfCanvas RestoreState() => Op("Q");

    public PdfCanvas SetFillColor(double r, double g, double b) => Op($"{N(r)} {N(g)} {N(b)} rg");
    public PdfCanvas SetStrokeColor(double r, double g, double b) => Op($"{N(r)} {N(g)} {N(b)} RG");
    public PdfCanvas SetLineWidth(double width) => Op($"{N(width)} w");

    public PdfCanvas MoveTo(double x, double y) => Op($"{N(x)} {N(y)} m");
    public PdfCanvas LineTo(double x, double y) => Op($"{N(x)} {N(y)} l");
    public PdfCanvas Rectangle(double x, double y, double width, double height) =>
        Op($"{N(x)} {N(y)} {N(width)} {N(height)} re");
    public PdfCanvas ClosePath() => Op("h");
    public PdfCanvas Stroke() => Op("S");
    public PdfCanvas Fill() => Op("f");
    public PdfCanvas FillAndStroke() => Op("B");

    /// <summary>Draws a single line of text with its baseline starting at (x, y).</summary>
    public PdfCanvas DrawText(PdfFont font, double fontSize, double x, double y, string text)
    {
        Fonts.Add(font);
        return Op($"BT /{font.ResourceName} {N(fontSize)} Tf {N(x)} {N(y)} Td {font.Encode(text)} Tj ET");
    }

    public PdfCanvas DrawImage(PdfImage image, double x, double y, double width, double height)
    {
        Images.Add(image);
        return Op($"q {N(width)} 0 0 {N(height)} {N(x)} {N(y)} cm /{image.ResourceName} Do Q");
    }

    internal byte[] ToBytes() => Encoding.ASCII.GetBytes(_content.ToString());

    private PdfCanvas Op(string op)
    {
        _content.Append(op).Append('\n');
        return this;
    }

    private static string N(double value) => PdfFormat.Number(value);
}
