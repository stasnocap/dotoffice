using System.Text;
using DotOffice.Core;
using DotOffice.Fonts;
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
}

internal sealed class Line
{
    public required List<Atom> Atoms { get; init; }
    public double Ascent { get; set; }
    public double Descent { get; set; }
    public bool EndsWithPageBreak { get; init; }
    public bool EndsWithLineBreak { get; init; }
    public bool IsLast { get; set; }

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
            }
        }

        var all = new StringBuilder();
        foreach (var s in segments) all.Append(s.Text);
        var text = all.ToString();
        var breaks = new LineBreakOpportunity[text.Length];
        LineBreaker.FindBreaks(text, breaks);

        var atoms = new List<Atom>();
        int offset = 0;
        foreach (var s in segments)
        {
            var before = offset < breaks.Length ? breaks[offset] : LineBreakOpportunity.None;
            switch (s.Kind)
            {
                case AtomKind.Tab:
                    atoms.Add(new Atom { Kind = AtomKind.Tab, Props = s.Props, Font = fonts.ForCharacter(s.Props, ' '), Size = s.Props.FontSize, BreakBefore = before, Link = s.Inline?.Link });
                    break;
                case AtomKind.Image:
                    var image = (ImageRun)s.Inline!;
                    atoms.Add(new Atom { Kind = AtomKind.Image, Props = s.Props, Image = image, Width = image.Width, BreakBefore = before, Link = image.Link });
                    break;
                case AtomKind.Marker:
                    // Glued to the preceding text so a bookmark never starts a line of its own.
                    atoms.Add(new Atom { Kind = AtomKind.Marker, Props = s.Props, Bookmark = ((BookmarkStart)s.Inline!).Name, BreakBefore = LineBreakOpportunity.None });
                    break;
                case AtomKind.LineBreak:
                case AtomKind.PageBreak:
                    atoms.Add(new Atom { Kind = s.Kind, Props = s.Props, Font = fonts.ForCharacter(s.Props, ' '), Size = s.Props.FontSize, BreakBefore = before });
                    break;
                default:
                    AddTextAtoms(atoms, s, text, offset, breaks);
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

    private void AddTextAtoms(List<Atom> atoms, Segment s, string all, int offset, LineBreakOpportunity[] breaks)
    {
        var props = s.Props;
        var field = s.Inline as FieldRun;
        double baseSize = props.FontSize;
        double size = props.VerticalPosition == VerticalPosition.Baseline ? baseSize : baseSize * 2 / 3;
        double rise = props.BaselineShift + props.VerticalPosition switch
        {
            VerticalPosition.Superscript => baseSize / 3,
            VerticalPosition.Subscript => -baseSize * 0.14,
            _ => 0,
        };

        int i = 0;
        var text = s.Text;
        while (i < text.Length)
        {
            int start = i;
            int cp = char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i];
            bool isSpace = text[i] == ' ';
            bool small = props.SmallCaps && char.IsLower(text[i]);
            var font = isSpace ? fonts.ForCharacter(props, ' ') : fonts.ForCharacterWithFallback(props, cp);
            i += char.IsSurrogatePair(text, start) ? 2 : 1;
            while (i < text.Length)
            {
                // A field result stays one atom: its text is substituted after pagination.
                if (field is null && breaks[offset + i] != LineBreakOpportunity.None) break;
                if (field is null && (text[i] == ' ') != isSpace) break;
                if (props.SmallCaps && char.IsLower(text[i]) != small) break;
                int next = char.IsSurrogatePair(text, i) ? char.ConvertToUtf32(text, i) : text[i];
                if (field is null && !isSpace && !char.IsLowSurrogate(text[i]) && !ReferenceEquals(fonts.ForCharacterWithFallback(props, next).Font, font.Font)
                    && !IsCombining(next))
                    break;
                i += char.IsSurrogatePair(text, i) ? 2 : 1;
            }
            var piece = text[start..i];
            double atomSize = small ? size * 0.8 : size;
            if (small) piece = piece.ToUpperInvariant();
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
                Width = font.Font.MeasureText(piece, atomSize) + props.CharacterSpacing * piece.Length,
                BreakBefore = offset + start < breaks.Length ? breaks[offset + start] : LineBreakOpportunity.None,
            });
        }
    }

    private static bool IsCombining(int cp) =>
        System.Globalization.CharUnicodeInfo.GetUnicodeCategory(cp) is System.Globalization.UnicodeCategory.NonSpacingMark
            or System.Globalization.UnicodeCategory.EnclosingMark or System.Globalization.UnicodeCategory.SpacingCombiningMark;
}

/// <summary>Geometry a paragraph's lines are filled into. All x values are relative to the text area's left edge.</summary>
internal readonly record struct LineFrame(double Width, double LeftIndent, double RightIndent, double FirstLineIndent,
    IReadOnlyList<TabStop> TabStops, double DefaultTabStop);

/// <summary>Greedy line filling over break opportunities, with tab stops, alignment and justification.</summary>
internal static class LineBuilder
{
    private const double Epsilon = 0.01;

    public static List<Line> Build(List<Atom> atoms, LineFrame frame, ParagraphAlignment alignment, bool justifyBeforeBreaks)
    {
        var lines = new List<Line>();
        int i = 0;
        while (i < atoms.Count)
        {
            bool first = lines.Count == 0;
            double startX = frame.LeftIndent + (first ? frame.FirstLineIndent : 0);
            double maxX = frame.Width - frame.RightIndent;
            int end = i;
            while (end < atoms.Count)
            {
                int wordEnd = end + 1;
                while (wordEnd < atoms.Count && atoms[wordEnd].BreakBefore == LineBreakOpportunity.None && !atoms[wordEnd - 1].IsBreak)
                    wordEnd++;
                double contentEnd = Position(atoms, i, wordEnd, startX, frame);
                if (contentEnd > maxX + Epsilon && end > i) break; // word does not fit: line ends before it
                if (contentEnd > maxX + Epsilon && end == i)
                {
                    // A single word wider than the line: break it between characters.
                    if (SplitToFit(atoms, i, wordEnd, startX, maxX, frame)) continue;
                }
                end = wordEnd;
                if (atoms[end - 1].IsBreak) break;
                if (end < atoms.Count && atoms[end].BreakBefore == LineBreakOpportunity.Mandatory) break;
            }
            if (end == i) end = i + 1; // always make progress

            var lineAtoms = atoms.GetRange(i, end - i);
            var last = lineAtoms[^1];
            lines.Add(new Line
            {
                Atoms = lineAtoms,
                EndsWithPageBreak = last.Kind == AtomKind.PageBreak,
                EndsWithLineBreak = last.Kind == AtomKind.LineBreak,
            });
            i = end;
        }

        // A trailing break leaves an empty line holding the paragraph mark.
        if (lines.Count == 0 || lines[^1].Atoms[^1].IsBreak)
            lines.Add(new Line { Atoms = [] });
        lines[^1].IsLast = true;

        for (int l = 0; l < lines.Count; l++)
            Finish(lines[l], frame, alignment, l == 0, justifyBeforeBreaks);
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
        atom.Font!.Font.MeasureText(text, atom.Size) + atom.Props.CharacterSpacing * text.Length;

    private static Atom Copy(Atom a, string text, LineBreakOpportunity breakBefore) => new()
    {
        Kind = a.Kind, Text = text, Props = a.Props, Font = a.Font, Size = a.Size, Rise = a.Rise, Field = a.Field, Link = a.Link,
        Width = Measure(a, text), BreakBefore = breakBefore,
    };

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