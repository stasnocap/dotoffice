using System.Text;

namespace DotOffice.Pdf;

/// <summary>Text rendering tweaks: synthetic bold (fill + stroke), synthetic italic (skew) and Tc spacing.</summary>
public readonly record struct TextStyle(bool FakeBold = false, bool FakeItalic = false, double CharacterSpacing = 0);

/// <summary>A piece of text whose first glyph starts at the absolute x position <see cref="X"/>.</summary>
public readonly record struct TextSegment(double X, string Text);

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

    /// <summary>Sets a dash pattern (on/off lengths); call with no arguments for solid lines.</summary>
    public PdfCanvas SetDash(params double[] pattern) =>
        Op($"[{string.Join(" ", pattern.Select(N))}] 0 d");

    /// <summary>0 = butt, 1 = round, 2 = projecting square.</summary>
    public PdfCanvas SetLineCap(int style) => Op($"{style} J");

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

    /// <summary>
    /// Draws several text segments sharing font, size and baseline in one text object. Gaps between
    /// a segment's natural end and the next segment's x are encoded as TJ adjustments, which lets
    /// layout position words exactly (e.g. for justification) without one text object per word.
    /// </summary>
    public PdfCanvas DrawTextSegments(PdfFont font, double fontSize, double y, IReadOnlyList<TextSegment> segments, TextStyle style = default)
    {
        if (segments.Count == 0) return this;
        Fonts.Add(font);
        bool wrap = style.FakeBold;
        if (wrap) Op("q");
        var sb = new StringBuilder("BT /").Append(font.ResourceName).Append(' ').Append(N(fontSize)).Append(" Tf ");
        if (style.FakeBold) sb.Append("2 Tr ").Append(N(fontSize * 0.03)).Append(" w ");
        if (style.CharacterSpacing != 0) sb.Append(N(style.CharacterSpacing)).Append(" Tc ");
        double x = segments[0].X;
        sb.Append(style.FakeItalic ? "1 0 0.2126 1 " : "1 0 0 1 ").Append(N(x)).Append(' ').Append(N(y)).Append(" Tm [");
        for (int i = 0; i < segments.Count; i++)
        {
            var segment = segments[i];
            double gap = segment.X - x;
            if (i > 0 && Math.Abs(gap) > 0.001) sb.Append(N(-gap * 1000 / fontSize)).Append(' ');
            sb.Append(font.Encode(segment.Text)).Append(' ');
            x = segment.X + font.MeasureText(segment.Text, fontSize) + style.CharacterSpacing * CountRunes(segment.Text);
        }
        sb.Append("] TJ ET");
        Op(sb.ToString());
        return wrap ? Op("Q") : this;
    }

    public PdfCanvas DrawImage(PdfImage image, double x, double y, double width, double height)
    {
        Images.Add(image);
        return Op($"q {N(width)} 0 0 {N(height)} {N(x)} {N(y)} cm /{image.ResourceName} Do Q");
    }

    private static int CountRunes(string text)
    {
        int count = 0;
        foreach (var _ in text.EnumerateRunes()) count++;
        return count;
    }

    internal byte[] ToBytes() => Encoding.ASCII.GetBytes(_content.ToString());

    private PdfCanvas Op(string op)
    {
        _content.Append(op).Append('\n');
        return this;
    }

    private static string N(double value) => PdfFormat.Number(value);
}
