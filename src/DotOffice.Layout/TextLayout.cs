using System.Text;
using DotOffice.Core;
using DotOffice.Fonts;
using DotOffice.Fonts.Shaping;
using DotOffice.Layout.Bidi;
using DotOffice.Layout.LineBreaking;

namespace DotOffice.Layout;

internal enum AtomKind { Text, Space, Tab, Image, LineBreak, PageBreak, Marker }

/// <summary>
/// The smallest unit of inline layout: a stretch of text in one font without break opportunities
/// inside it, a run of spaces, a tab, an image or a break. Words are sequences of atoms between
/// break opportunities.
/// </summary>
internal sealed class Atom
{
    public AtomKind Kind { get; init; }
    public string Text { get; init; } = "";
    public required RunProperties Props { get; init; }
    public ResolvedFont? Font { get; init; }
    public double Size { get; init; }

    /// <summary>Baseline raise in points (superscript, w:position).</summary>
    public double Rise { get; init; }

    public double Width { get; init; }
    public ImageRun? Image { get; init; }
    public FieldRun? Field { get; init; }
    public Hyperlink? Link { get; init; }

    /// <summary>Bookmark name of a <see cref="AtomKind.Marker"/> atom.</summary>
    public string? Bookmark { get; init; }

    /// <summary>The note referenced by this atom (its mark, or a marker when the mark is custom text).</summary>
    public Note? Note { get; init; }
    public TabStop? Tab { get; set; }

    /// <summary>Part of the list label text (not its suffix).</summary>
    public bool IsLabel { get; init; }
    public LineBreakOpportunity BreakBefore { get; set; }

    // Set by positioning.
    public double X { get; set; }
    public double Advance { get; set; }

    public FontMetrics Metrics => Font is null ? default : FontMetrics.Of(Font.Font, Size);

    public int SpaceCount => Kind == AtomKind.Space ? Text.Length : 0;

    public bool IsBreak => Kind is AtomKind.LineBreak or AtomKind.PageBreak;

    /// <summary>A <see cref="AtomKind.PageBreak"/> atom that only ends the column.</summary>
    public bool IsColumnBreak { get; init; }

    /// <summary>Bidi embedding level; odd levels are right to left.</summary>
    public byte Level { get; init; }

    public bool RightToLeft => (Level & 1) == 1;

    /// <summary>Shaped glyphs in logical order, or null when the text maps to glyphs one to one.</summary>
    public ShapedGlyph[]? Glyphs { get; init; }
}

internal sealed class Line
{
    public required List<Atom> Atoms { get; init; }
    public double Ascent { get; set; }
    public double Descent { get; set; }
    public bool EndsWithPageBreak { get; init; }
    public bool EndsWithColumnBreak { get; init; }
    public bool EndsWithLineBreak { get; init; }
    public bool IsLast { get; set; }

    /// <summary>Extra space above the line, pushing it below an obstacle such as a floating image.</summary>
    public double Gap { get; init; }

    /// <summary>Atoms that are drawn (trailing spaces and breaks excluded).</summary>
    public int VisibleCount { get; set; }
}

/// <summary>Converts paragraph inlines into atoms.</summary>
internal sealed class Itemizer(FontSelector fonts)
{
    private readonly record struct Segment(AtomKind Kind, string Text, RunProperties Props, Inline? Inline, bool IsLabel = false);

    /// <summary>Builds atoms for a paragraph; <paramref name="fieldText"/> supplies text for page-number fields.</summary>
    public List<Atom> Itemize(Paragraph paragraph, Func<FieldRun, string> fieldText)
    {
        var segments = new List<Segment>();
        if (paragraph.ListLabel is { } label && !label.Properties.Hidden)
        {
            segments.Add(new Segment(AtomKind.Text, Transform(label.Text, label.Properties), label.Properties, null, IsLabel: true));
            if (label.Suffix == ListLabelSuffix.Tab) segments.Add(new Segment(AtomKind.Tab, "\t", label.Properties, null));
            else if (label.Suffix == ListLabelSuffix.Space) segments.Add(new Segment(AtomKind.Text, " ", label.Properties, null));
        }
        foreach (var inline in paragraph.Inlines)
        {
            if (inline.Properties.Hidden) continue;
            switch (inline)
            {
                case TextRun t:
                    AddText(segments, Transform(t.Text, t.Properties), t.Properties, t);
                    break;
                case Tab:
                    segments.Add(new Segment(AtomKind.Tab, "\t", inline.Properties, inline));
                    break;
                case ImageRun:
                    segments.Add(new Segment(AtomKind.Image, "￼", inline.Properties, inline));
                    break;
                case Break { Kind: BreakKind.Line }:
                    segments.Add(new Segment(AtomKind.LineBreak, "\n", inline.Properties, inline));
                    break;
                case Break:
                    segments.Add(new Segment(AtomKind.PageBreak, "\f", inline.Properties, inline));
                    break;
                case FieldRun f:
                    segments.Add(new Segment(AtomKind.Text, fieldText(f), inline.Properties, inline));
                    break;
                case BookmarkStart:
                    segments.Add(new Segment(AtomKind.Marker, "", inline.Properties, inline));
                    break;
                case NoteReference note:
                    segments.Add(note.Mark.Length > 0
                        ? new Segment(AtomKind.Text, Transform(note.Mark, note.Properties), note.Properties, note)
                        : new Segment(AtomKind.Marker, "", note.Properties, note));
                    break;
                case NoteMark { Note.Mark.Length: > 0 } mark:
                    segments.Add(new Segment(AtomKind.Text, Transform(mark.Note.Mark, mark.Properties), mark.Properties, mark));
                    break;
            }
        }

        var all = new StringBuilder();
        foreach (var s in segments) all.Append(s.Text);
        var text = all.ToString();
        var breaks = new LineBreakOpportunity[text.Length];
        LineBreaker.FindBreaks(text, breaks);
        byte paragraphLevel = paragraph.Properties.RightToLeft ? (byte)1 : (byte)0;
        var levels = BidiLevels(segments, text, paragraphLevel);
        byte Level(int index) => levels is null ? (byte)0 : index < levels.Length ? levels[index] : paragraphLevel;

        var atoms = new List<Atom>();
        int offset = 0;
        foreach (var s in segments)
        {
            var before = offset < breaks.Length ? breaks[offset] : LineBreakOpportunity.None;
            switch (s.Kind)
            {
                case AtomKind.Tab:
                    atoms.Add(new Atom
                    {
                        Kind = AtomKind.Tab, Props = s.Props, Font = fonts.ForCharacter(s.Props, ' '), Size = s.Props.FontSize, BreakBefore = before,
                        Link = s.Inline?.Link, Level = Level(offset),
                    });
                    break;
                case AtomKind.Image:
                    var image = (ImageRun)s.Inline!;
                    atoms.Add(new Atom
                    {
                        Kind = AtomKind.Image, Props = s.Props, Image = image, Width = image.Width, BreakBefore = before, Link = image.Link,
                        Level = Level(offset),
                    });
                    break;
                case AtomKind.Marker:
                    // Glued to the preceding text so a bookmark never starts a line of its own.
                    atoms.Add(new Atom
                    {
                        Kind = AtomKind.Marker, Props = s.Props, BreakBefore = LineBreakOpportunity.None,
                        Bookmark = (s.Inline as BookmarkStart)?.Name, Note = (s.Inline as NoteReference)?.Note, Level = Level(offset),
                    });
                    break;
                case AtomKind.LineBreak:
                case AtomKind.PageBreak:
                    atoms.Add(new Atom
                    {
                        Kind = s.Kind, Props = s.Props, Font = fonts.ForCharacter(s.Props, ' '), Size = s.Props.FontSize, BreakBefore = before,
                        IsColumnBreak = s.Inline is Break { Kind: BreakKind.Column }, Level = paragraphLevel,
                    });
                    break;
                default:
                    AddTextAtoms(atoms, s, offset, breaks, Level);
                    break;
            }
            offset += s.Text.Length;
        }
        return atoms;
    }

    private static void AddText(List<Segment> segments, string text, RunProperties props, Inline source)
    {
        int start = 0;
        for (int i = 0; i <= text.Length; i++)
        {
            char c = i < text.Length ? text[i] : '\0';
            if (i < text.Length && c is not ('\t' or '\n' or '\r' or '\v' or '\f')) continue;
            if (i > start) segments.Add(new Segment(AtomKind.Text, text[start..i], props, source));
            if (c == '\t') segments.Add(new Segment(AtomKind.Tab, "\t", props, source));
            else if (c is '\n' or '\v' || (c == '\r' && (i + 1 >= text.Length || text[i + 1] != '\n')))
                segments.Add(new Segment(AtomKind.LineBreak, "\n", props, source));
            else if (c == '\f') segments.Add(new Segment(AtomKind.PageBreak, "\f", props, source));
            start = i + 1;
        }
    }

    private static string Transform(string text, RunProperties props) =>
        props.Caps ? text.ToUpperInvariant() : text;

    /// <summary>
    /// Bidi embedding levels per UTF-16 index of the paragraph text, or null when everything is at level 0
    /// (a left-to-right paragraph without right-to-left text). Breaks count as whitespace, not paragraph
    /// separators; neutral characters of runs marked right-to-left (w:rtl) are treated as right-to-left.
    /// </summary>
    private static byte[]? BidiLevels(List<Segment> segments, string text, byte paragraphLevel)
    {
        bool any = paragraphLevel == 1 || segments.Exists(s => s.Props.RightToLeft && s.Text.Length > 0);
        if (!any)
            foreach (var rune in text.EnumerateRunes())
                if (BidiAlgorithm.GetClass(rune.Value) is BidiClass.R or BidiClass.AL or BidiClass.AN or BidiClass.RLE or BidiClass.RLO
                    or BidiClass.RLI or BidiClass.FSI) { any = true; break; }
        if (!any) return null;

        var classes = new BidiClass[text.Length];
        var codePoints = new int[text.Length];
        int offset = 0;
        foreach (var s in segments)
        {
            for (int i = 0; i < s.Text.Length; i++)
            {
                int cp = char.IsSurrogatePair(s.Text, i) ? char.ConvertToUtf32(s.Text, i) : s.Text[i];
                var cls = s.Kind is AtomKind.LineBreak or AtomKind.PageBreak ? BidiClass.WS
                    : char.IsLowSurrogate(s.Text[i]) ? BidiClass.BN : BidiAlgorithm.GetClass(cp);
                if (s.Props.RightToLeft && cls is BidiClass.WS or BidiClass.ON) cls = BidiClass.R;
                classes[offset + i] = cls;
                codePoints[offset + i] = cp;
            }
            offset += s.Text.Length;
        }
        var levels = BidiAlgorithm.ResolveLevels(classes, codePoints, paragraphLevel);
        // The low half of a surrogate pair goes with its high half.
        for (int i = 1; i < text.Length; i++)
            if (char.IsLowSurrogate(text[i])) levels[i] = levels[i - 1];
        return levels;
    }

    private void AddTextAtoms(List<Atom> atoms, Segment s, int offset, LineBreakOpportunity[] breaks, Func<int, byte> level)
    {
        var props = s.Props;
        var field = s.Inline as FieldRun;

        int i = 0;
        var text = s.Text;
        while (i < text.Length)
        {
            int start = i;
            int cp = char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i];
            bool isSpace = text[i] == ' ';
            bool small = props.SmallCaps && char.IsLower(text[i]);
            bool complex = FontSelector.UsesComplexScript(props, cp);
            byte atomLevel = level(offset + start);
            var font = isSpace ? fonts.ForCharacter(props, ' ') : fonts.ForCharacterWithFallback(props, cp);
            if (TextShaper.IsUnsupportedComplexScript(cp))
                fonts.WarnOnce("layout.complex-script", "Indic and similar scripts need reordering that is not implemented; their text may render incorrectly.");
            i += char.IsSurrogatePair(text, start) ? 2 : 1;
            while (i < text.Length)
            {
                // A field result stays one atom: its text is substituted after pagination.
                if (field is null && breaks[offset + i] != LineBreakOpportunity.None) break;
                if (field is null && (text[i] == ' ') != isSpace) break;
                if (props.SmallCaps && char.IsLower(text[i]) != small) break;
                if (char.IsLowSurrogate(text[i])) { i++; continue; }
                int next = char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i];
                if (field is null && level(offset + i) != atomLevel) break;
                if (field is null && !isSpace && !IsCombining(next)
                    && (FontSelector.UsesComplexScript(props, next) != complex || !ReferenceEquals(fonts.ForCharacterWithFallback(props, next).Font, font.Font)))
                    break;
                i += char.IsSurrogatePair(text, i) ? 2 : 1;
            }

            var piece = StripBidiControls(text[start..i]);
            if (piece.Length == 0) continue;
            double baseSize = FontSelector.SizeFor(props, cp);
            double size = props.VerticalPosition == VerticalPosition.Baseline ? baseSize : baseSize * 2 / 3;
            double rise = props.BaselineShift + props.VerticalPosition switch
            {
                VerticalPosition.Superscript => baseSize / 3,
                VerticalPosition.Subscript => -baseSize * 0.14,
                _ => 0,
            };
            double atomSize = small ? size * 0.8 : size;
            if (small) piece = piece.ToUpperInvariant();
            bool rtl = (atomLevel & 1) == 1;
            var (width, glyphs) = isSpace ? (Measure(font, piece, atomSize, props, rtl, shape: false), null) : Shape(font, piece, atomSize, props, rtl);
            atoms.Add(new Atom
            {
                Kind = isSpace ? AtomKind.Space : AtomKind.Text,
                IsLabel = s.IsLabel,
                Text = piece,
                Props = props,
                Font = font,
                Size = atomSize,
                Rise = rise,
                Field = field,
                Link = s.Inline?.Link,
                Note = (s.Inline as NoteReference)?.Note,
                Level = atomLevel,
                Glyphs = glyphs,
                Width = width,
                BreakBefore = offset + start < breaks.Length ? breaks[offset + start] : LineBreakOpportunity.None,
            });
        }
    }

    /// <summary>
    /// Shapes text when it needs it (complex script, marks, kerning, or right-to-left so its glyphs can be drawn
    /// in reverse with mirrored brackets); returns its width and glyphs (null when unshaped). The glyphs keep the
    /// logical characters as their text.
    /// </summary>
    internal static (double Width, ShapedGlyph[]? Glyphs) Shape(ResolvedFont font, string text, double size, RunProperties props, bool rtl)
    {
        bool kerning = props.KerningMinSize is { } min && size >= min;
        if (!rtl && !TextShaper.NeedsShaping(text, kerning)) return (Measure(font, text, size, props, rtl, shape: false), null);
        var glyphs = TextShaper.Shape(font.Font, rtl ? Mirror(text) : text, rtl, kerning);
        // Mirrored glyphs still stand for the logical character in extracted text.
        if (rtl)
            for (int i = 0; i < glyphs.Length; i++) glyphs[i] = glyphs[i] with { Text = Mirror(glyphs[i].Text) };
        double units = 0;
        int advancing = 0;
        foreach (var g in glyphs)
        {
            units += g.Advance;
            if (g.Advance != 0) advancing++;
        }
        return (units * size / font.Font.UnitsPerEm + props.CharacterSpacing * advancing, glyphs);
    }

    /// <summary>Width of text in a font; shaped when <paramref name="shape"/> is set and the text needs it.</summary>
    internal static double Measure(ResolvedFont font, string text, double size, RunProperties props, bool rtl, bool shape = true)
    {
        if (shape || rtl) return Shape(font, text, size, props, rtl).Width;
        int runes = 0;
        foreach (var _ in text.EnumerateRunes()) runes++;
        return font.Font.MeasureText(text, size) + props.CharacterSpacing * runes;
    }

    /// <summary>Bidi formatting characters steer the algorithm but are never drawn.</summary>
    private static string StripBidiControls(string text) =>
        text.AsSpan().IndexOfAny("\u200E\u200F\u061C\u202A\u202B\u202C\u202D\u202E\u2066\u2067\u2068\u2069") < 0
            ? text
            : string.Concat(text.Where(c => c is not ('\u200E' or '\u200F' or '\u061C' or >= '\u202A' and <= '\u202E' or >= '\u2066' and <= '\u2069')));

    /// <summary>Replaces characters with their mirrored forms (brackets…) for right-to-left display (L4).</summary>
    private static string Mirror(string text)
    {
        StringBuilder? sb = null;
        for (int i = 0; i < text.Length; i++)
        {
            if (char.IsSurrogate(text[i])) continue;
            int mirrored = BidiAlgorithm.Mirror(text[i]);
            if (mirrored == text[i] || mirrored > 0xFFFF) continue;
            sb ??= new StringBuilder(text);
            sb[i] = (char)mirrored;
        }
        return sb?.ToString() ?? text;
    }

    private static bool IsCombining(int cp) =>
        System.Globalization.CharUnicodeInfo.GetUnicodeCategory(cp) is System.Globalization.UnicodeCategory.NonSpacingMark
            or System.Globalization.UnicodeCategory.EnclosingMark or System.Globalization.UnicodeCategory.SpacingCombiningMark;
}

/// <summary>Geometry a paragraph's lines are filled into. All x values are relative to the text area's left edge.</summary>
/// <remarks>For right-to-left paragraphs lines are filled as if left to right and mirrored afterwards, so
/// <see cref="LeftIndent"/> and left alignment refer to the start (right) side.</remarks>
internal readonly record struct LineFrame(double Width, double LeftIndent, double RightIndent, double FirstLineIndent,
    IReadOnlyList<TabStop> TabStops, double DefaultTabStop, bool RightToLeft = false);

/// <summary>
/// Supplies per-line geometry while a paragraph is broken into lines (text wrapping around floats):
/// <see cref="Next"/> is asked before each line, <see cref="Done"/> is told about each finished line.
/// </summary>
internal interface ILineSpace
{
    /// <summary>Frame for the next line and the gap to leave above it.</summary>
    (LineFrame Frame, double Gap) Next(LineFrame frame, bool first);

    void Done(Line line);
}

/// <summary>Greedy line filling over break opportunities, with tab stops, alignment and justification.</summary>
internal static class LineBuilder
{
    private const double Epsilon = 0.01;

    public static List<Line> Build(List<Atom> atoms, LineFrame frame, ParagraphAlignment alignment, bool justifyBeforeBreaks,
        ILineSpace? space = null)
    {
        var lines = new List<Line>();
        int i = 0;
        while (i < atoms.Count)
        {
            bool first = lines.Count == 0;
            var (f, gap) = space?.Next(frame, first) ?? (frame, 0);
            double startX = f.LeftIndent + (first ? f.FirstLineIndent : 0);
            double maxX = f.Width - f.RightIndent;
            int end = i;
            while (end < atoms.Count)
            {
                int wordEnd = end + 1;
                while (wordEnd < atoms.Count && atoms[wordEnd].BreakBefore == LineBreakOpportunity.None && !atoms[wordEnd - 1].IsBreak)
                    wordEnd++;
                double contentEnd = Position(atoms, i, wordEnd, startX, f);
                if (contentEnd > maxX + Epsilon && end > i) break; // word does not fit: line ends before it
                if (contentEnd > maxX + Epsilon && end == i)
                {
                    // A single word wider than the line: break it between characters.
                    if (SplitToFit(atoms, i, wordEnd, startX, maxX, f)) continue;
                }
                end = wordEnd;
                if (atoms[end - 1].IsBreak) break;
                if (end < atoms.Count && atoms[end].BreakBefore == LineBreakOpportunity.Mandatory) break;
            }
            if (end == i) end = i + 1; // always make progress

            var lineAtoms = atoms.GetRange(i, end - i);
            var last = lineAtoms[^1];
            var line = new Line
            {
                Atoms = lineAtoms,
                EndsWithPageBreak = last.Kind == AtomKind.PageBreak,
                EndsWithColumnBreak = last.IsColumnBreak,
                Gap = gap,
                IsLast = end == atoms.Count && !last.IsBreak,
            };
            lines.Add(line);
            i = end;
            Finish(line, f, alignment, first, justifyBeforeBreaks);
            space?.Done(line);
        }

        // A trailing break leaves an empty line holding the paragraph mark.
        if (lines.Count == 0 || lines[^1].Atoms[^1].IsBreak)
        {
            var (f, gap) = space?.Next(frame, lines.Count == 0) ?? (frame, 0);
            var line = new Line { Atoms = [], Gap = gap, IsLast = true };
            lines.Add(line);
            Finish(line, f, alignment, lines.Count == 1, justifyBeforeBreaks);
            space?.Done(line);
        }
        return lines;
    }

    /// <summary>Assigns X/Advance to atoms [from, to) and returns the end of visible content.</summary>
    private static double Position(List<Atom> atoms, int from, int to, double startX, LineFrame frame)
    {
        double x = startX, contentEnd = startX;
        for (int k = from; k < to; k++)
        {
            var atom = atoms[k];
            atom.X = x;
            if (atom.Kind == AtomKind.Tab)
            {
                var stop = NextStop(x, frame);
                atom.Tab = stop;
                double segment = 0, beforeDecimal = -1;
                if (stop.Alignment is TabAlignment.Right or TabAlignment.Center or TabAlignment.Decimal)
                {
                    int lastVisible = to - 1;
                    while (lastVisible > k && atoms[lastVisible].Kind is AtomKind.Space or AtomKind.LineBreak or AtomKind.PageBreak) lastVisible--;
                    for (int m = k + 1; m <= lastVisible && atoms[m].Kind != AtomKind.Tab; m++)
                    {
                        var a = atoms[m];
                        if (stop.Alignment == TabAlignment.Decimal && beforeDecimal < 0 && a.Kind == AtomKind.Text
                            && a.Text.IndexOfAny(['.', ',']) is var dot and >= 0)
                            beforeDecimal = segment + a.Font!.Font.MeasureText(a.Text[..dot], a.Size);
                        segment += a.Width;
                    }
                }
                double target = stop.Alignment switch
                {
                    TabAlignment.Right => stop.Position - segment,
                    TabAlignment.Center => stop.Position - segment / 2,
                    TabAlignment.Decimal => stop.Position - (beforeDecimal >= 0 ? beforeDecimal : segment),
                    _ => stop.Position,
                };
                atom.Advance = Math.Max(0, target - x);
                x += atom.Advance;
                contentEnd = x;
            }
            else
            {
                atom.Advance = atom.Width;
                x += atom.Width;
                if (atom.Kind is AtomKind.Text or AtomKind.Image) contentEnd = x;
            }
        }
        return contentEnd;
    }

    private static TabStop NextStop(double x, LineFrame frame)
    {
        foreach (var stop in frame.TabStops)
            if (stop.Alignment != TabAlignment.Bar && stop.Position > x + Epsilon)
            {
                // A hanging indent acts as an implicit stop before later custom stops.
                if (frame.FirstLineIndent < 0 && frame.LeftIndent > x + Epsilon && frame.LeftIndent < stop.Position)
                    return new TabStop(frame.LeftIndent, TabAlignment.Left);
                return stop;
            }
        if (frame.FirstLineIndent < 0 && frame.LeftIndent > x + Epsilon)
            return new TabStop(frame.LeftIndent, TabAlignment.Left);
        double interval = frame.DefaultTabStop > 0 ? frame.DefaultTabStop : 36;
        double next = (Math.Floor((x + Epsilon) / interval) + 1) * interval;
        return new TabStop(next, TabAlignment.Left);
    }

    /// <summary>Splits the first oversized text atom of a word so that a prefix fits on the line.</summary>
    private static bool SplitToFit(List<Atom> atoms, int from, int to, double startX, double maxX, LineFrame frame)
    {
        double x = startX;
        for (int k = from; k < to; k++)
        {
            var atom = atoms[k];
            double width = atom.Kind == AtomKind.Tab ? 0 : atom.Width;
            if (atom.Kind == AtomKind.Text && x + width > maxX + Epsilon)
            {
                var text = atom.Text;
                int fit = 0;
                for (int n = 1; n < text.Length; n++)
                {
                    if (char.IsLowSurrogate(text[n])) continue;
                    if (x + Measure(atom, text[..n]) > maxX + Epsilon) break;
                    fit = n;
                }
                if (fit == 0 && k == from)
                {
                    fit = char.IsSurrogatePair(text, 0) ? 2 : 1; // at least one character per line
                    if (fit >= text.Length) return false;
                }
                if (fit == 0)
                {
                    atom.BreakBefore = LineBreakOpportunity.Allowed;
                    return true;
                }
                atoms[k] = Copy(atom, text[..fit], atom.BreakBefore);
                atoms.Insert(k + 1, Copy(atom, text[fit..], LineBreakOpportunity.Allowed));
                return true;
            }
            x += width;
        }
        _ = frame;
        return false;
    }

    private static double Measure(Atom atom, string text) =>
        Itemizer.Measure(atom.Font!, text, atom.Size, atom.Props, atom.RightToLeft, shape: atom.Glyphs is not null);

    private static Atom Copy(Atom a, string text, LineBreakOpportunity breakBefore)
    {
        var (width, glyphs) = a.Glyphs is null
            ? (Itemizer.Measure(a.Font!, text, a.Size, a.Props, rtl: false, shape: false), null)
            : Itemizer.Shape(a.Font!, text, a.Size, a.Props, a.RightToLeft);
        return new Atom
        {
            Kind = a.Kind, Text = text, Props = a.Props, Font = a.Font, Size = a.Size, Rise = a.Rise, Field = a.Field, Link = a.Link, Note = a.Note,
            Level = a.Level, Glyphs = glyphs, Width = width, BreakBefore = breakBefore,
        };
    }

    private static void Finish(Line line, LineFrame frame, ParagraphAlignment alignment, bool first, bool justifyBeforeBreaks)
    {
        var atoms = line.Atoms;
        double startX = frame.LeftIndent + (first ? frame.FirstLineIndent : 0);
        double maxX = frame.Width - frame.RightIndent;
        double contentEnd = Position(atoms, 0, atoms.Count, startX, frame);

        int visible = atoms.Count;
        while (visible > 0 && atoms[visible - 1].Kind is AtomKind.Space or AtomKind.LineBreak or AtomKind.PageBreak or AtomKind.Marker) visible--;
        line.VisibleCount = visible;

        double extra = maxX - contentEnd;
        bool justify = alignment == ParagraphAlignment.Distribute
                       || (alignment == ParagraphAlignment.Justify && !line.IsLast && (!line.EndsWithLineBreak || justifyBeforeBreaks)
                           && !line.EndsWithPageBreak);
        if (justify && extra > 0)
        {
            int lastTab = -1;
            for (int k = 0; k < visible; k++) if (atoms[k].Kind == AtomKind.Tab) lastTab = k;
            int spaces = 0;
            for (int k = lastTab + 1; k < visible; k++) spaces += atoms[k].SpaceCount;
            if (spaces > 0)
            {
                double per = extra / spaces, shift = 0;
                for (int k = lastTab + 1; k < atoms.Count; k++)
                {
                    atoms[k].X += shift;
                    if (k < visible && atoms[k].Kind == AtomKind.Space)
                    {
                        double add = per * atoms[k].SpaceCount;
                        atoms[k].Advance += add;
                        shift += add;
                    }
                }
            }
        }
        else if (extra > 0 && alignment is ParagraphAlignment.Center or ParagraphAlignment.Right)
        {
            double offset = alignment == ParagraphAlignment.Center ? extra / 2 : extra;
            foreach (var atom in atoms) atom.X += offset;
        }

        if (frame.RightToLeft || atoms.Exists(a => a.Level > 0)) Reorder(line, frame);

        double ascent = 0, descent = 0;
        foreach (var atom in atoms)
        {
            if (atom.Kind == AtomKind.Image)
            {
                ascent = Math.Max(ascent, atom.Image!.Height);
                continue;
            }
            if (atom.Font is null) continue;
            var m = atom.Metrics;
            ascent = Math.Max(ascent, m.Ascent + atom.Rise);
            descent = Math.Max(descent, m.Descent - atom.Rise);
        }
        line.Ascent = ascent;
        line.Descent = descent;
    }

    /// <summary>
    /// Puts a line's atoms in visual order (bidi rule L2 on atom levels) and positions them left to right
    /// inside the span the logical layout gave the line; right-to-left paragraphs mirror that span first.
    /// The drawn atoms of <see cref="Line.Atoms"/> end up in visual order.
    /// </summary>
    private static void Reorder(Line line, LineFrame frame)
    {
        var atoms = line.Atoms;
        int visible = line.VisibleCount;
        if (visible == 0) return;
        double start = atoms[0].X, end = atoms[visible - 1].X + atoms[visible - 1].Advance;
        double x = frame.RightToLeft ? frame.Width - end : start;
        var levels = new byte[visible];
        for (int k = 0; k < visible; k++) levels[k] = atoms[k].Level;
        var visual = BidiAlgorithm.VisualOrder(levels).Select(k => atoms[k]).ToList();
        for (int k = 0; k < visible; k++)
        {
            atoms[k] = visual[k];
            atoms[k].X = x;
            x += atoms[k].Advance;
        }
        for (int k = visible; k < atoms.Count; k++) atoms[k].X = x;
    }

    /// <summary>Minimum (longest word) and maximum (no wrapping) content widths, for auto table layout.</summary>
    public static (double Min, double Max) Intrinsic(List<Atom> atoms, double defaultTab)
    {
        double min = 0, max = 0, word = 0, line = 0, pendingSpace = 0;
        foreach (var atom in atoms)
        {
            if (atom.BreakBefore != LineBreakOpportunity.None)
            {
                min = Math.Max(min, word);
                word = 0;
            }
            switch (atom.Kind)
            {
                case AtomKind.Space:
                    pendingSpace += atom.Width;
                    continue;
                case AtomKind.LineBreak:
                case AtomKind.PageBreak:
                    max = Math.Max(max, line);
                    line = pendingSpace = 0;
                    min = Math.Max(min, word);
                    word = 0;
                    continue;
            }
            double w = atom.Kind == AtomKind.Tab ? defaultTab : atom.Width;
            line += pendingSpace + w;
            pendingSpace = 0;
            word += w;
        }
        return (Math.Max(min, word), Math.Max(max, line));
    }
}