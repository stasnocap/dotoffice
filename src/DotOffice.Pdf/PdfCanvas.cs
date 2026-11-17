using System.Text;

namespace DotOffice.Pdf;

/// <summary>
/// Text rendering tweaks: synthetic bold (fill + stroke), synthetic italic (skew) and Tc spacing. With <c>AsPaths</c> the glyphs are filled as outlines rather than written as text, so that they cannot be selected or
/// found (a shadow of text that is written once already); a font without TrueType outlines then draws nothing.
/// </summary>
public readonly record struct TextStyle(bool FakeBold = false, bool FakeItalic = false, double CharacterSpacing = 0, bool AsPaths = false);

/// <summary>A piece of text whose first glyph starts at the absolute x position <see cref="X"/>.</summary>
public readonly record struct TextSegment(double X, string Text);

/// <summary>
/// A glyph drawn at absolute position (<see cref="X"/>, baseline + <see cref="Rise"/>); <see cref="Text"/> is
/// what it stands for in the extracted text (may be empty).
/// </summary>
public readonly record struct PositionedGlyph(ushort GlyphId, string Text, double X, double Rise);

/// <summary>
/// Builds a page content stream. Coordinates are in points with the origin at the
/// bottom-left corner of the page (native PDF space).
/// </summary>
public sealed class PdfCanvas
{
    private readonly StringBuilder _content = new();
    private readonly Stack<double> _savedCharSpacing = new();
    private double _charSpacing;

    internal HashSet<PdfFont> Fonts { get; } = [];
    internal HashSet<PdfImage> Images { get; } = [];
    internal HashSet<PdfShading> Shadings { get; } = [];

    /// <summary>The glyph outlines drawn as forms (see <see cref="FillGlyphOutlines"/>), by font and glyph; resource names Gl1, Gl2, ...</summary>
    internal Dictionary<(DotOffice.Fonts.TrueTypeFont Font, ushort Glyph), (string Name, byte[] Content)> GlyphForms { get; } = [];

    /// <summary>The opacities (in percent) set with <see cref="SetAlpha"/>; each becomes a graphics state in the page resources.</summary>
    internal SortedSet<int> Alphas { get; } = [];

    /// <summary>Makes everything painted from now on (until the state is restored) partly transparent: 1 is opaque, 0 invisible.</summary>
    public PdfCanvas SetAlpha(double opacity)
    {
        int percent = (int)Math.Round(Math.Clamp(opacity, 0, 1) * 100);
        Alphas.Add(percent);
        return Op($"/GSa{percent} gs");
    }

    public PdfCanvas SaveState()
    {
        _savedCharSpacing.Push(_charSpacing);
        return Op("q");
    }

    public PdfCanvas RestoreState()
    {
        if (_savedCharSpacing.TryPop(out var spacing)) _charSpacing = spacing;
        return Op("Q");
    }

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
    public PdfCanvas CurveTo(double x1, double y1, double x2, double y2, double x3, double y3) =>
        Op($"{N(x1)} {N(y1)} {N(x2)} {N(y2)} {N(x3)} {N(y3)} c");
    public PdfCanvas ClosePath() => Op("h");

    /// <summary>Intersects the clip with the current path (use inside <see cref="SaveState"/>/<see cref="RestoreState"/>).</summary>
    public PdfCanvas Clip() => Op("W n");

    /// <summary>Like <see cref="Clip"/>, but by the even-odd rule: a ring inside another is cut out.</summary>
    public PdfCanvas ClipEvenOdd() => Op("W* n");

    /// <summary>Multiplies the current transformation matrix; points map as x' = a·x + c·y + e, y' = b·x + d·y + f.</summary>
    public PdfCanvas Transform(double a, double b, double c, double d, double e, double f) =>
        Op($"{N(a)} {N(b)} {N(c)} {N(d)} {N(e)} {N(f)} cm");

    /// <summary>Paints a shading over the whole current clip.</summary>
    public PdfCanvas PaintShading(PdfShading shading)
    {
        Shadings.Add(shading);
        return Op($"/{shading.ResourceName} sh");
    }
    public PdfCanvas Stroke() => Op("S");
    public PdfCanvas Fill() => Op("f");

    /// <summary>Fills by the even-odd rule: a ring inside another is left empty.</summary>
    public PdfCanvas FillEvenOdd() => Op("f*");
    public PdfCanvas FillAndStroke() => Op("B");

    /// <summary>Fills by the even-odd rule, then strokes.</summary>
    public PdfCanvas FillAndStrokeEvenOdd() => Op("B*");

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
        if (style.AsPaths)
        {
            var glyphs = new List<PositionedGlyph>();
            double scale = fontSize / font.Font.UnitsPerEm;
            foreach (var segment in segments)
            {
                double gx = segment.X;
                foreach (var rune in segment.Text.EnumerateRunes())
                {
                    ushort id = font.Font.GetGlyphId(rune.Value);
                    glyphs.Add(new PositionedGlyph(id, rune.ToString(), gx, 0));
                    gx += font.Font.GetAdvanceWidth(id) * scale + style.CharacterSpacing;
                }
            }
            return FillGlyphOutlines(font.Font, fontSize, y, glyphs, style.FakeItalic);
        }
        Fonts.Add(font);
        // Tc persists in the graphics state, so it is tracked and only emitted when it changes.
        if (style.CharacterSpacing != _charSpacing)
        {
            Op($"{N(style.CharacterSpacing)} Tc");
            _charSpacing = style.CharacterSpacing;
        }
        bool wrap = style.FakeBold;
        if (wrap) SaveState();
        var sb = new StringBuilder("BT /").Append(font.ResourceName).Append(' ').Append(N(fontSize)).Append(" Tf ");
        if (style.FakeBold) sb.Append("2 Tr ").Append(N(fontSize * 0.03)).Append(" w ");
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
        return wrap ? RestoreState() : this;
    }

    /// <summary>
    /// Fills glyphs as their outlines. Each glyph's outline (quadratic TrueType curves drawn as cubics, in font units) is a form drawn in the
    /// current fill color, so that many copies of the same text (a blurred shadow) cost one placement per glyph.
    /// </summary>
    private PdfCanvas FillGlyphOutlines(DotOffice.Fonts.TrueTypeFont font, double fontSize, double y, IReadOnlyList<PositionedGlyph> glyphs, bool italic)
    {
        double scale = fontSize / font.UnitsPerEm, slant = italic ? 0.2126 : 0;
        var sb = new StringBuilder();
        foreach (var g in glyphs)
        {
            if (!GlyphForms.TryGetValue((font, g.GlyphId), out var form))
            {
                var outline = GlyphOutline(font, g.GlyphId);
                if (outline.Length == 0) continue;
                GlyphForms[(font, g.GlyphId)] = form = ($"Gl{GlyphForms.Count + 1}", outline);
            }
            if (form.Content.Length == 0) continue;
            sb.Append("q ").Append(N(scale)).Append(" 0 ").Append(N(slant * scale)).Append(' ').Append(N(scale)).Append(' ')
              .Append(N(g.X)).Append(' ').Append(N(y + g.Rise)).Append(" cm /").Append(form.Name).Append(" Do Q ");
        }
        return sb.Length > 0 ? Op(sb.ToString().TrimEnd()) : this;
    }

    /// <summary>A glyph's outline as a filled path in font units; empty for a blank glyph.</summary>
    private static byte[] GlyphOutline(DotOffice.Fonts.TrueTypeFont font, ushort glyphId)
    {
        var sb = new StringBuilder();
        string P(double ux, double uy) => $"{N(ux)} {N(uy)}";
        foreach (var contour in font.GetGlyphOutline(glyphId))
        {
            int n = contour.Count;
            if (n < 2) continue;
            // Start on a point on the curve (or between two control points when there is none).
            int first = Enumerable.Range(0, n).FirstOrDefault(k => contour[k].OnCurve, -1);
            (double X, double Y) start = first >= 0 ? (contour[first].X, contour[first].Y)
                : ((contour[0].X + contour[1].X) / 2, (contour[0].Y + contour[1].Y) / 2);
            int from = first >= 0 ? first : 0;
            sb.Append(P(start.X, start.Y)).Append(" m ");
            var current = start;
            (double X, double Y)? control = null;
            for (int k = 1; k <= n; k++)
            {
                var pt = contour[(from + k) % n];
                bool closing = k == n;
                if (pt.OnCurve || closing && first >= 0)
                {
                    var end = closing ? start : (pt.X, pt.Y);
                    if (control is { } q) Quad(q, end); else sb.Append(P(end.X, end.Y)).Append(" l ");
                    control = null;
                }
                else
                {
                    if (control is { } q) Quad(q, ((q.X + pt.X) / 2, (q.Y + pt.Y) / 2));
                    control = (pt.X, pt.Y);
                    if (closing) Quad(control.Value, start);
                }
            }
            sb.Append("h ");

            void Quad((double X, double Y) q, (double X, double Y) end)
            {
                sb.Append(P(current.X + 2.0 / 3 * (q.X - current.X), current.Y + 2.0 / 3 * (q.Y - current.Y))).Append(' ')
                  .Append(P(end.X + 2.0 / 3 * (q.X - end.X), end.Y + 2.0 / 3 * (q.Y - end.Y))).Append(' ')
                  .Append(P(end.X, end.Y)).Append(" c ");
                current = end;
            }
        }
        return sb.Length == 0 ? [] : Encoding.ASCII.GetBytes(sb.Append('f').ToString());
    }

    /// <summary>
    /// Draws shaped glyphs (already in visual order) on the baseline <paramref name="y"/>. Consecutive glyphs on
    /// the same rise share one TJ array whose adjustments place each glyph exactly.
    /// </summary>
    public PdfCanvas DrawGlyphs(PdfFont font, double fontSize, double y, IReadOnlyList<PositionedGlyph> glyphs, TextStyle style = default)
    {
        if (glyphs.Count == 0) return this;
        if (style.AsPaths) return FillGlyphOutlines(font.Font, fontSize, y, glyphs, style.FakeItalic);
        Fonts.Add(font);
        if (_charSpacing != 0)
        {
            Op("0 Tc"); // glyph positions already include any character spacing
            _charSpacing = 0;
        }
        bool wrap = style.FakeBold;
        if (wrap) SaveState();
        var sb = new StringBuilder("BT /").Append(font.ResourceName).Append(' ').Append(N(fontSize)).Append(" Tf ");
        if (style.FakeBold) sb.Append("2 Tr ").Append(N(fontSize * 0.03)).Append(" w ");
        string matrix = style.FakeItalic ? "1 0 0.2126 1 " : "1 0 0 1 ";
        double scale = fontSize / font.Font.UnitsPerEm;
        for (int i = 0; i < glyphs.Count;)
        {
            double rise = glyphs[i].Rise, x = glyphs[i].X;
            sb.Append(matrix).Append(N(x)).Append(' ').Append(N(y + rise)).Append(" Tm [");
            for (; i < glyphs.Count && glyphs[i].Rise == rise; i++)
            {
                var g = glyphs[i];
                double gap = g.X - x;
                if (Math.Abs(gap) > 0.001) sb.Append(N(-gap * 1000 / fontSize)).Append(' ');
                sb.Append('<').Append(font.EncodeGlyph(g.GlyphId, g.Text)).Append("> ");
                x = g.X + font.Font.GetAdvanceWidth(g.GlyphId) * scale;
            }
            sb.Append("] TJ ");
        }
        sb.Append("ET");
        Op(sb.ToString());
        return wrap ? RestoreState() : this;
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
