using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>
/// Lays out the worksheets of a <see cref="Workbook"/> into pages: each sheet's print range is split into bands of
/// columns and rows that fit a page, and every cell is drawn into its band with its fill, borders and text.
/// </summary>
public sealed partial class SheetLayoutEngine
{
    private const double Epsilon = 0.01;

    /// <summary>Horizontal space between a cell edge and its text.</summary>
    private const double CellPadding = 2;

    /// <summary>Width of one indent level.</summary>
    private const double IndentWidth = 6.75;

    /// <summary>Excel's row height for text: 15 pt for an 11 pt font.</summary>
    private const double LineHeightRatio = 15.0 / 11.0;

    private const double MaxRowHeight = 409.5;

    private const int MaxColumns = 16384;

    private const double MinScale = 0.1;

    private static readonly Color GridColor = new(0xA6, 0xA6, 0xA6);

    private readonly FontSelector _fonts;
    private readonly Itemizer _itemizer;

    /// <summary>Creates an engine that measures text with the fonts of <paramref name="options"/>.</summary>
    public SheetLayoutEngine(LayoutOptions options)
    {
        _fonts = new FontSelector(options.Fonts, options.Warnings);
        _itemizer = new Itemizer(_fonts);
    }

    /// <summary>Lays out every visible sheet of the workbook; an empty workbook yields one blank page.</summary>
    public LayoutDocument Layout(Workbook workbook, CancellationToken cancellationToken = default)
    {
        var document = new LayoutDocument { Title = workbook.Title, Author = workbook.Author };
        var records = new List<PageRecord>();
        foreach (var sheet in workbook.Sheets)
        {
            if (sheet.Visibility != SheetVisibility.Visible) continue;
            new SheetPass(this, sheet, document, records, cancellationToken).Run();
        }
        if (document.Pages.Count == 0) document.Pages.Add(new LayoutPage(612, 792));

        // Page numbers continue across sheets, so headers and footers wait until the total is known.
        // A sheet may restart the numbering at its own first page number; the pages after it continue from there.
        int number = 1;
        for (int i = 0; i < records.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var setup = records[i].Sheet.PageSetup;
            if (records[i].IndexInSheet == 0 && setup.UseFirstPageNumber) number = setup.FirstPageNumber;
            EmitHeaderFooter(records[i], number++, records.Count, workbook);
        }
        return document;
    }

    /// <summary>
    /// Where an object lies once its group fraction (<see cref="SheetObject.Frame"/>) is taken out of the anchor's frame and the rotated groups
    /// it belongs to have swung it about their centers.
    /// </summary>
    internal static (double X, double Y, double Width, double Height) PlaceInFrame(SheetObject image, double x, double y, double width, double height)
    {
        var f = image.Frame;
        var rect = (X: x + f.X * width, Y: y + f.Y * height, Width: f.Width * width, Height: f.Height * height);
        if (image.GroupTurns.Count == 0) return rect;
        // The rotated groups swing the object about their centers (innermost first); its own turn, already counting theirs, is applied later.
        double cx = rect.X + rect.Width / 2, cy = rect.Y + rect.Height / 2;
        for (int i = image.GroupTurns.Count - 1; i >= 0; i--)
        {
            var (gf, degrees) = (image.GroupTurns[i].Frame, image.GroupTurns[i].Rotation);
            double gx = x + (gf.X + gf.Width / 2) * width, gy = y + (gf.Y + gf.Height / 2) * height;
            double angle = degrees * Math.PI / 180, dx = cx - gx, dy = cy - gy;
            (cx, cy) = (gx + dx * Math.Cos(angle) + dy * Math.Sin(angle), gy - dx * Math.Sin(angle) + dy * Math.Cos(angle));
        }
        return (cx - rect.Width / 2, cy - rect.Height / 2, rect.Width, rect.Height);
    }

    /// <summary>A page of a sheet, kept until headers and footers can be drawn.</summary>
    private sealed record PageRecord(LayoutPage Page, Worksheet Sheet, int IndexInSheet, double Scale = 1);

    // ------------------------------------------------------------------ headers and footers

    private static readonly RunProperties HeaderFont = new() { FontFamily = "Calibri", FontSize = 11 };

    private void EmitHeaderFooter(PageRecord record, int number, int total, Workbook workbook)
    {
        var setup = record.Sheet.PageSetup;
        bool first = record.IndexInSheet == 0;
        static (HeaderFooterParts Parts, string Variant) Pick(HeaderFooterText text, bool first, int number) =>
            first && text.First is { } f ? (f, "FIRST") : number % 2 == 0 && text.Even is { } e ? (e, "EVEN") : (text.Odd, "");
        var (header, headerVariant) = Pick(setup.Header, first, number);
        var (footer, footerVariant) = Pick(setup.Footer, first, number);
        if (header.IsEmpty && footer.IsEmpty) return;
        var context = new FieldContext(number, total, record.Sheet.Name, workbook.Title ?? "", setup.HeaderFooterScaleWithDocument ? record.Scale : 1);
        var page = record.Page;
        double left = setup.Margins.Left, right = page.Width - setup.Margins.Right;

        void Part(string text, double anchorLeft, double anchorRight, int alignment, bool isHeader, string variant)
        {
            if (string.IsNullOrEmpty(text)) return;
            // A picture (&G) sits above the section's text, as wide and high as its drawing says (scaled with the sheet).
            HeaderFooterPicture? picture = null;
            if (text.Contains("&G", StringComparison.Ordinal))
            {
                // The first-page and even-page variants have their own pictures; without one the default picture is used.
                string key = string.Concat("LCR"[alignment], isHeader ? 'H' : 'F');
                if (!setup.HeaderFooterPictures.TryGetValue(key + variant, out picture)) setup.HeaderFooterPictures.TryGetValue(key, out picture);
                text = text.Replace("&G", "", StringComparison.Ordinal);
            }
            var lines = text.Replace("\r\n", "\n").Split('\n').Select(l => BuildHeaderLine(l, context)).Where(l => l is not null).Select(l => l!).ToList();
            if (lines.Count == 0 && picture is null) return;
            double pitch(HeaderLine l) => l.Size * 1.2;
            double pictureWidth = (picture?.Width ?? 0) * context.Scale, pictureHeight = (picture?.Height ?? 0) * context.Scale;
            double height = lines.Sum(pitch) + pictureHeight;
            double top = isHeader ? setup.HeaderDistance : page.Height - setup.FooterDistance - height;
            if (picture is not null)
            {
                double px = alignment switch
                {
                    0 => anchorLeft,
                    1 => (anchorLeft + anchorRight) / 2 - pictureWidth / 2,
                    _ => anchorRight - pictureWidth,
                };
                page.Items.Add(new ImageItem(px, top, pictureWidth, pictureHeight, picture.Image));
                top += pictureHeight;
            }
            foreach (var line in lines)
            {
                double x0 = alignment switch
                {
                    0 => anchorLeft,
                    1 => (anchorLeft + anchorRight) / 2 - line.Width / 2,
                    _ => anchorRight - line.Width,
                };
                var metrics = line.Metrics;
                double baseline = top + (pitch(line) - (metrics.Ascent + metrics.Descent)) / 2 + metrics.Ascent;
                var items = new List<LayoutItem>();
                ParagraphLayouter.DrawLine(line.Line, x0, baseline, items);
                page.Items.AddRange(items);
                top += pitch(line);
            }
        }

        foreach (var (parts, isHeader, variant) in new[] { (header, true, headerVariant), (footer, false, footerVariant) })
        {
            Part(parts.Left, left, right, 0, isHeader, variant);
            Part(parts.Center, left, right, 1, isHeader, variant);
            Part(parts.Right, left, right, 2, isHeader, variant);
        }
    }

    private readonly record struct FieldContext(int Page, int Pages, string Sheet, string File, double Scale = 1);

    private sealed record HeaderLine(Line Line, double Width, double Size, FontMetrics Metrics);

    /// <summary>Turns one line of header/footer text (with its <c>&amp;</c> codes) into drawable atoms.</summary>
    private HeaderLine? BuildHeaderLine(string text, FieldContext context)
    {
        var runs = new List<(string Text, RunProperties Props)>();
        var buffer = new System.Text.StringBuilder();
        var props = HeaderFont;
        double maxSize = 0;

        void Flush()
        {
            if (buffer.Length == 0) return;
            // With scaleWithDoc the text is as large as the page scale makes it.
            var scaled = context.Scale == 1 ? props : props with { FontSize = props.FontSize * context.Scale };
            runs.Add((buffer.ToString(), scaled));
            maxSize = Math.Max(maxSize, scaled.FontSize);
            buffer.Clear();
        }

        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c != '&' || i + 1 >= text.Length)
            {
                buffer.Append(c);
                continue;
            }
            char code = text[++i];
            switch (code)
            {
                case '&': buffer.Append('&'); break;
                case 'P': buffer.Append(context.Page.ToString(System.Globalization.CultureInfo.InvariantCulture)); break;
                case 'N': buffer.Append(context.Pages.ToString(System.Globalization.CultureInfo.InvariantCulture)); break;
                case 'A': buffer.Append(context.Sheet); break;
                case 'F': buffer.Append(context.File); break;
                case 'D': buffer.Append(DateTime.Today.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture)); break;
                case 'T': buffer.Append(DateTime.Now.ToString("h:mm tt", System.Globalization.CultureInfo.InvariantCulture)); break;
                case 'B': Flush(); props = props with { Bold = !props.Bold }; break;
                case 'I': Flush(); props = props with { Italic = !props.Italic }; break;
                case 'U': Flush(); props = props with { Underline = props.Underline == UnderlineStyle.None ? UnderlineStyle.Single : UnderlineStyle.None }; break;
                case 'E': Flush(); props = props with { Underline = UnderlineStyle.Double }; break;
                case 'S': Flush(); props = props with { Strike = !props.Strike }; break;
                case '"':
                {
                    int end = text.IndexOf('"', i + 1);
                    if (end < 0) end = text.Length;
                    var spec = text.Substring(i + 1, end - i - 1).Split(',');
                    Flush();
                    var family = spec[0].Trim();
                    if (family.Length > 0 && family != "-") props = props with { FontFamily = family };
                    if (spec.Length > 1)
                    {
                        var style = spec[1];
                        props = props with
                        {
                            Bold = style.Contains("Bold", StringComparison.OrdinalIgnoreCase),
                            Italic = style.Contains("Italic", StringComparison.OrdinalIgnoreCase),
                        };
                    }
                    i = end;
                    break;
                }
                case 'K' when IsHex(text, i + 1, 6):
                    Flush();
                    props = props with { Color = new Color(Convert.ToByte(text.Substring(i + 1, 2), 16), Convert.ToByte(text.Substring(i + 3, 2), 16), Convert.ToByte(text.Substring(i + 5, 2), 16)) };
                    i += 6;
                    break;
                case >= '0' and <= '9':
                {
                    int digits = 1;
                    if (i + 1 < text.Length && char.IsAsciiDigit(text[i + 1])) digits = 2;
                    double size = double.Parse(text.AsSpan(i, digits), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture);
                    Flush();
                    if (size > 0) props = props with { FontSize = size };
                    i += digits - 1;
                    break;
                }
                default: break; // &G, &X, &Y, &Z … are ignored
            }
        }
        Flush();
        if (runs.Count == 0) return null;

        var paragraph = new Paragraph { MarkProperties = runs[0].Props };
        foreach (var (runText, runProps) in runs) paragraph.Inlines.Add(new TextRun(runText, runProps));
        var atoms = _itemizer.Itemize(paragraph, _ => "");
        var lines = LineBuilder.Build(atoms, new LineFrame(1e7, 0, 0, 0, [], 36), ParagraphAlignment.Left, false, null);
        if (lines.Count == 0) return null;
        var line = lines[0];
        var metrics = FontMetrics.Of(_fonts.ForCharacter(runs[0].Props, '0').Font, maxSize);
        return new HeaderLine(line, LineWidth(line), maxSize, metrics);
    }

    private static bool IsHex(string text, int start, int count) =>
        start + count <= text.Length && text.AsSpan(start, count).ToString().All(char.IsAsciiHexDigit);

    // ------------------------------------------------------------------ text

    /// <summary>The lines of one cell's text, ready to be drawn with <see cref="ParagraphLayouter.DrawLine"/>.</summary>
    private sealed class TextBlock
    {
        public List<Line> Lines { get; } = [];

        /// <summary>Per line: its height, and the ascent and descent of its tallest font (Excel sizes each line by its largest run).</summary>
        public List<(double Pitch, double Ascent, double Descent)> Metrics { get; } = [];

        public double Width { get; set; }
        public double Height => Metrics.Sum(m => m.Pitch);
    }

    private static double LineWidth(Line line)
    {
        double width = 0;
        for (int k = 0; k < line.VisibleCount && k < line.Atoms.Count; k++)
            width = Math.Max(width, line.Atoms[k].X + line.Atoms[k].Advance);
        return width;
    }

    private static ParagraphAlignment AlignmentOf(Cell cell) => cell.Style.HorizontalAlignment switch
    {
        HorizontalAlignment.Center => ParagraphAlignment.Center,
        HorizontalAlignment.Right => ParagraphAlignment.Right,
        HorizontalAlignment.Justify => ParagraphAlignment.Justify,
        HorizontalAlignment.General => cell.Kind switch
        {
            CellValueKind.Number => ParagraphAlignment.Right,
            CellValueKind.Bool or CellValueKind.Error => ParagraphAlignment.Center,
            _ => ParagraphAlignment.Left,
        },
        _ => ParagraphAlignment.Left,
    };

    /// <summary>
    /// The paragraphs (each a list of formatted runs) of a cell's text. Rich-text runs are kept when the text is the cell's own;
    /// for wrapping cells a newline starts a new paragraph, otherwise newlines are dropped.
    /// </summary>
    private static List<List<TextRun>> Paragraphs(Cell cell, string text, bool wrap)
    {
        var runs = cell.Runs is { Count: > 0 } r && text == cell.Text
            ? r.Select(x => (x.Text, x.Font)).ToList()
            : [(text, cell.Style.Font)];
        var paragraphs = new List<List<TextRun>> { new() };
        if (cell.Style.Stacked)
        {
            // Stacked text: every character is a line of its own.
            paragraphs.Clear();
            foreach (var (runText, font) in runs)
                foreach (var rune in runText.EnumerateRunes())
                    if (rune.Value is not ('\r' or '\n')) paragraphs.Add([new TextRun(rune.ToString(), font)]);
            if (paragraphs.Count == 0) paragraphs.Add([]);
            return paragraphs;
        }
        foreach (var (runText, font) in runs)
        {
            var normalized = runText.Replace("\r\n", "\n").Replace('\r', '\n');
            if (!wrap) normalized = normalized.Replace("\n", "");
            var parts = normalized.Split('\n');
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0) paragraphs.Add([]);
                if (parts[i].Length > 0) paragraphs[^1].Add(new TextRun(parts[i], font));
            }
        }
        return paragraphs;
    }

    /// <summary>Breaks a cell's text into lines; <paramref name="wrapWidth"/> is the text width for wrapping cells.</summary>
    private TextBlock BuildBlock(Cell cell, string text, double wrapWidth, double leftIndent = 0, double firstLineIndent = 0)
    {
        var font = cell.Style.Font;
        bool wrap = cell.Style.Wrap && cell.Kind == CellValueKind.Text;
        var paragraphs = Paragraphs(cell, text, wrap);

        var block = new TextBlock();
        (double Pitch, double Ascent, double Descent) Measure(RunProperties props)
        {
            var m = FontMetrics.Of(_fonts.ForCharacter(props, '0').Font, props.FontSize);
            return (Math.Max(props.FontSize * LineHeightRatio, 1), m.Ascent, m.Descent);
        }
        var firstRun = paragraphs.SelectMany(p => p).Select(r => r.Properties).FirstOrDefault() ?? font;

        var alignment = wrap ? AlignmentOf(cell) : ParagraphAlignment.Left;
        var frame = new LineFrame(wrap ? Math.Max(wrapWidth, 1) : 1e7, leftIndent, 0, firstLineIndent, [], 36);
        foreach (var runs in paragraphs)
        {
            var tallest = runs.Select(r => r.Properties).DefaultIfEmpty(firstRun).MaxBy(p => p.FontSize)!;
            var lines = new List<Line>();
            if (runs.Count > 0)
            {
                var paragraph = new Paragraph { MarkProperties = font };
                foreach (var run in runs) paragraph.Inlines.Add(run);
                var atoms = _itemizer.Itemize(paragraph, _ => "");
                lines = LineBuilder.Build(atoms, frame, alignment, false, null);
            }
            if (lines.Count == 0) lines.Add(new Line { Atoms = [] });
            // A wrapped paragraph's lines are sized by the tallest run on them; the paragraph's tallest run is a close upper bound.
            foreach (var line in lines)
            {
                var props = line.Atoms.Take(line.VisibleCount).Where(a => a.Font is not null).Select(a => a.Props).DefaultIfEmpty(tallest).MaxBy(p => p.FontSize)!;
                block.Lines.Add(line);
                block.Metrics.Add(Measure(props));
            }
        }
        foreach (var line in block.Lines) block.Width = Math.Max(block.Width, LineWidth(line));
        return block;
    }

    private static bool IsRotated(CellStyle style) => style.Rotation != 0 || style.Stacked;

    /// <summary>
    /// The text of a rotated or stacked cell, laid out unrotated, and the size of its bounding box once rotated.
    /// Rotated text is only wrapped when the cell wraps; Excel measures the line against the cell's width (divided by the cosine
    /// of the angle) and, for upright vertical text, against its height. When <paramref name="height"/> is null (row autofit) such
    /// vertical text is not wrapped, so the row grows to the longest line instead.
    /// </summary>
    private (TextBlock Block, double Width, double Height) RotatedBlock(Cell cell, double width, double? height)
    {
        var style = cell.Style;
        double radians = style.Rotation * Math.PI / 180, cos = Math.Abs(Math.Cos(radians)), sin = Math.Abs(Math.Sin(radians));
        double wrapLength = 1e7;
        if (style.Wrap && !style.Stacked)
        {
            if (cos > 0.01) wrapLength = Math.Max((width - 2 * CellPadding) / cos, 1);
            else if (height is { } h) wrapLength = Math.Max(h - 2 * CellPadding, 1);
        }
        var block = BuildBlock(cell, cell.Text, wrapLength);
        double w = block.Width, h2 = block.Height;
        return style.Stacked ? (block, w, h2) : (block, w * cos + h2 * sin, w * sin + h2 * cos);
    }

    /// <summary>A copy of a text atom with different text (used to clip text to a cell).</summary>
    private Atom Trimmed(Atom atom, string text, double x) => new()
    {
        Kind = atom.Kind, Text = text, Props = atom.Props, Font = atom.Font, Size = atom.Size, Rise = atom.Rise,
        Width = Itemizer.Measure(atom.Font!, text, atom.Size, atom.Props, false),
        Advance = Itemizer.Measure(atom.Font!, text, atom.Size, atom.Props, false), X = x,
    };

    /// <summary>Draws the part of a line that lies between <paramref name="clipLeft"/> and <paramref name="clipRight"/>.</summary>
    private void EmitLine(Line line, double x0, double baseline, double clipLeft, double clipRight, List<LayoutItem> items)
    {
        var kept = new List<Atom>();
        for (int k = 0; k < line.VisibleCount && k < line.Atoms.Count; k++)
        {
            var a = line.Atoms[k];
            double start = x0 + a.X, end = start + a.Advance;
            if (end <= clipLeft + Epsilon || start >= clipRight - Epsilon)
            {
                if (a.Advance > Epsilon) continue;
            }
            if (start >= clipLeft - Epsilon && end <= clipRight + Epsilon)
            {
                kept.Add(a);
                continue;
            }
            if (a.Kind != AtomKind.Text || a.Glyphs is not null || a.Font is null) continue;

            // Partly outside: drop characters from the clipped sides.
            string text = a.Text;
            double x = a.X;
            while (text.Length > 0 && x0 + x + Itemizer.Measure(a.Font, text, a.Size, a.Props, false) > clipRight + Epsilon)
                text = text[..(text.Length - (text.Length > 1 && char.IsLowSurrogate(text[^1]) ? 2 : 1))];
            while (text.Length > 0 && x0 + x < clipLeft - Epsilon)
            {
                int n = char.IsHighSurrogate(text[0]) && text.Length > 1 ? 2 : 1;
                x += Itemizer.Measure(a.Font, text[..n], a.Size, a.Props, false);
                text = text[n..];
            }
            if (text.Length > 0) kept.Add(Trimmed(a, text, x));
        }
        if (kept.Count == 0) return;
        ParagraphLayouter.DrawLine(new Line { Atoms = kept, VisibleCount = kept.Count }, x0, baseline, items);
    }

    // ------------------------------------------------------------------ pages

    private static LayoutItem Transform(LayoutItem item, double dx, double dy, double scale) => item switch
    {
        TextItem t => t with { X = dx + t.X * scale, Baseline = dy + t.Baseline * scale, FontSize = t.FontSize * scale, CharacterSpacing = t.CharacterSpacing * scale },
        LineItem l => l with { X1 = dx + l.X1 * scale, X2 = dx + l.X2 * scale, Y1 = dy + l.Y1 * scale, Y2 = dy + l.Y2 * scale, Width = l.Width * scale },
        PushItem p => p with
        {
            Clip = p.Clip is { } c ? (dx + c.X * scale, dy + c.Y * scale, c.Width * scale, c.Height * scale) : null,
            CenterX = dx + p.CenterX * scale, CenterY = dy + p.CenterY * scale,
            ClipOutline = p.ClipOutline?.Select(q => (dx + q.X * scale, dy + q.Y * scale)).ToList(),
        },
        GradientItem g => g with { X = dx + g.X * scale, Y = dy + g.Y * scale, Width = g.Width * scale, Height = g.Height * scale, Clip = g.Clip?.Select(q => (dx + q.X * scale, dy + q.Y * scale)).ToList() },
        ImageItem i => i with { X = dx + i.X * scale, Y = dy + i.Y * scale, Width = i.Width * scale, Height = i.Height * scale, Clip = i.Clip?.Select(q => (dx + q.X * scale, dy + q.Y * scale)).ToList() },
        PolygonItem p => p with { Points = p.Points.Select(q => (dx + q.X * scale, dy + q.Y * scale)).ToList() },
        EllipseItem e => e with { X = dx + e.X * scale, Y = dy + e.Y * scale, Width = e.Width * scale, Height = e.Height * scale },
        RectItem r => r with { X = dx + r.X * scale, Y = dy + r.Y * scale, Width = r.Width * scale, Height = r.Height * scale },
        _ => item,
    };

    /// <summary>The state of laying out one worksheet.</summary>
    private sealed class SheetPass
    {
        private readonly SheetLayoutEngine _engine;
        private readonly Worksheet _sheet;
        private readonly LayoutDocument _document;
        private readonly List<PageRecord> _records;
        private readonly CancellationToken _ct;
        private int _sheetPages;
        private readonly SheetPageSetup _setup;

        private CellRange _range;
        private readonly Dictionary<int, double> _rowHeights = [];
        private readonly Dictionary<(int Row, int Column), TextBlock> _blocks = [];
        private readonly Dictionary<(int Row, int Column), CellRange> _anchors = [];

        public SheetPass(SheetLayoutEngine engine, Worksheet sheet, LayoutDocument document, List<PageRecord> records, CancellationToken ct)
        {
            _engine = engine;
            _records = records;
            _sheet = sheet;
            _document = document;
            _ct = ct;
            _setup = sheet.PageSetup;
        }

        public void Run()
        {
            foreach (var merge in _sheet.Merges) _anchors[(merge.FirstRow, merge.FirstColumn)] = merge;
            // Each print area is printed as its own page set (Excel starts a new page for every area).
            if (_setup.PrintAreas.Count == 0) RunRange(null);
            else foreach (var area in _setup.PrintAreas) RunRange(area);
            if (_setup.Comments != CommentPrinting.None && _sheet.Comments.Count > 0) EmitCommentPages();
        }

        /// <summary>
        /// Lists the comments on pages after the sheet, in cell order, as "Cell A1 (Author)" followed by the text. The list is laid out at
        /// 100 % on the sheet's paper and margins and flows from page to page.
        /// </summary>
        private void EmitCommentPages()
        {
            double left = _setup.Margins.Left, width = Math.Max(_setup.PageWidth - _setup.Margins.Left - _setup.Margins.Right, 1);
            double top = _setup.Margins.Top, bottom = _setup.PageHeight - _setup.Margins.Bottom;
            var items = new List<LayoutItem>();
            double y = top;
            void FinishPage()
            {
                var page = new LayoutPage(_setup.PageWidth, _setup.PageHeight);
                foreach (var item in items) page.Items.Add(item);
                _document.Pages.Add(page);
                _records.Add(new PageRecord(page, _sheet, _sheetPages++));
                items.Clear();
                y = top;
            }
            void Write(Cell cell)
            {
                var block = _engine.BuildBlock(cell, cell.Text, width);
                for (int i = 0; i < block.Lines.Count; i++)
                {
                    var (pitch, ascent, descent) = block.Metrics[i];
                    if (y + pitch > bottom + Epsilon && y > top + Epsilon) FinishPage();
                    if (block.Lines[i].Atoms.Count > 0)
                        _engine.EmitLine(block.Lines[i], left, y + (pitch - (ascent + descent)) / 2 + ascent, left, left + width, items);
                    y += pitch;
                }
            }
            var normal = CellStyle.Default with { Wrap = true };
            var bold = normal with { Font = normal.Font with { Bold = true } };
            foreach (var comment in _sheet.Comments)
            {
                _ct.ThrowIfCancellationRequested();
                Write(new Cell { Text = comment.Author is null ? $"Cell {comment.Reference}" : $"Cell {comment.Reference} ({comment.Author})", Kind = CellValueKind.Text, Style = bold });
                if (comment.Text.Trim().Length > 0) Write(new Cell { Text = comment.Text.Trim(), Kind = CellValueKind.Text, Style = normal });
                y += 8;
            }
            if (items.Count > 0) FinishPage();
        }

        private void RunRange(CellRange? area)
        {
            if (FindRange(area) is not { } range) return;
            _range = range;
            _rowHeights.Clear();
            _blocks.Clear();

            var columns = Enumerable.Range(_range.FirstColumn, _range.ColumnCount).Where(c => _sheet.ColumnWidth(c) > 0).ToList();
            ComputeRowHeights();
            if (_setup.TitleRows is { } outside && outside.From < _range.FirstRow)
                ComputeRowHeights(outside.From, Math.Min(outside.To, _range.FirstRow - 1));
            var rows = Enumerable.Range(_range.FirstRow, _range.RowCount).Where(r => RowHeight(r) > 0).ToList();
            if (columns.Count == 0 || rows.Count == 0) return;

            // Print titles that lie before the print range are printed in front of every page, as Excel does; those after it are not.
            var titleColumns = _setup.TitleColumns is { } tc
                ? Enumerable.Range(tc.From, Math.Max(Math.Min(tc.To, _range.FirstColumn - 1) - tc.From + 1, 0)).Where(c => _sheet.ColumnWidth(c) > 0)
                    .Concat(columns.Where(c => c >= tc.From && c <= tc.To)).ToList()
                : [];
            var titleRows = _setup.TitleRows is { } tr
                ? Enumerable.Range(tr.From, Math.Max(Math.Min(tr.To, _range.FirstRow - 1) - tr.From + 1, 0)).Where(r => RowHeight(r) > 0)
                    .Concat(rows.Where(r => r >= tr.From && r <= tr.To)).ToList()
                : [];

            double pageWidth = _setup.PageWidth - _setup.Margins.Left - _setup.Margins.Right;
            double pageHeight = _setup.PageHeight - _setup.Margins.Top - _setup.Margins.Bottom;
            double scale = _setup.Scale;
            var columnBreaks = new SortedSet<int>(_setup.ColumnBreaks);
            var rowBreaks = new SortedSet<int>(_setup.RowBreaks);
            if (_setup.FitToWidth is { } fitWidth && _setup.FitToHeight is { } fitHeight)
            {
                // Fit to page only shrinks, ignores manual breaks, and is re-checked against the real band count.
                columnBreaks.Clear();
                rowBreaks.Clear();
                double totalWidth = columns.Sum(c => _sheet.ColumnWidth(c)), totalHeight = rows.Sum(RowHeight);
                scale = 1;
                if (fitWidth > 0 && totalWidth > 0) scale = Math.Min(scale, pageWidth * fitWidth / totalWidth);
                if (fitHeight > 0 && totalHeight > 0) scale = Math.Min(scale, pageHeight * fitHeight / totalHeight);
                scale = Math.Max(scale, MinScale);
                for (int attempt = 0; attempt < 60 && scale > MinScale; attempt++)
                {
                    int across = Bands(columns, c => _sheet.ColumnWidth(c), pageWidth / scale, columnBreaks, titleColumns).Count;
                    int down = Bands(rows, RowHeight, pageHeight / scale, rowBreaks, titleRows).Count;
                    if ((fitWidth == 0 || across <= fitWidth) && (fitHeight == 0 || down <= fitHeight)) break;
                    scale = Math.Max(scale * 0.97, MinScale);
                }
            }
            var columnBands = Bands(columns, c => _sheet.ColumnWidth(c), pageWidth / scale, columnBreaks, titleColumns);
            var rowBands = Bands(rows, RowHeight, pageHeight / scale, rowBreaks, titleRows);

            foreach (var (columnBand, rowBand) in PageOrder(columnBands, rowBands))
            {
                _ct.ThrowIfCancellationRequested();
                EmitPage(columnBand, rowBand, scale);
            }
        }

        private double RowHeight(int row) => _sheet.HiddenRows.Contains(row) ? 0 : _rowHeights.GetValueOrDefault(row, _sheet.DefaultRowHeight);

        // ---- range

        private CellRange? FindRange(CellRange? printArea)
        {
            int r0 = int.MaxValue, c0 = int.MaxValue, r1 = -1, c1 = -1;
            foreach (var (row, column) in _sheet.Cells.Keys)
            {
                r0 = Math.Min(r0, row);
                c0 = Math.Min(c0, column);
                r1 = Math.Max(r1, row);
                c1 = Math.Max(c1, column);
            }
            // Excel prints pictures even outside the cells, so they belong to the used range.
            foreach (var image in Objects())
            {
                var (ix, iy, iw, ih) = ImageRectFromA1(image);
                if (IsEmpty(image, iw, ih)) continue;
                int firstColumn = ColumnAt(ix), lastColumn = ColumnAt(ix + iw - Epsilon), firstRow = RowAt(iy), lastRow = RowAt(iy + ih - Epsilon);
                if (r1 < 0) (r0, c0) = (firstRow, firstColumn);
                (r0, c0, r1, c1) = (Math.Min(r0, firstRow), Math.Min(c0, firstColumn), Math.Max(r1, lastRow), Math.Max(c1, lastColumn));
            }
            foreach (var merge in _sheet.Merges)
            {
                if (r1 < 0) continue;
                r1 = Math.Max(r1, merge.LastRow);
                c1 = Math.Max(c1, merge.LastColumn);
            }
            if (printArea is { } area)
            {
                // A whole-column or whole-row area would print a million blank rows; stop at the last used cell.
                int lastRow = area.RowCount > 100_000 && r1 >= 0 ? Math.Min(area.LastRow, r1) : area.LastRow;
                int lastColumn = area.ColumnCount > 1_000 && c1 >= 0 ? Math.Min(area.LastColumn, c1) : area.LastColumn;
                return new CellRange(area.FirstRow, area.FirstColumn, Math.Max(lastRow, area.FirstRow), Math.Max(lastColumn, area.FirstColumn));
            }
            return r1 < 0 ? null : ExtendForOverflow(new CellRange(r0, c0, r1, c1));
        }

        /// <summary>Excel prints text that spills out of the used range, so the range grows to hold it.</summary>
        private CellRange ExtendForOverflow(CellRange range)
        {
            int first = range.FirstColumn, last = range.LastColumn;
            foreach (var ((row, column), cell) in _sheet.Cells)
            {
                if (cell.Kind != CellValueKind.Text || cell.Style.Wrap || cell.Text.Length == 0 || IsRotated(cell.Style)) continue;
                double own = _sheet.ColumnWidth(column);
                // Cheap filter: skip texts that cannot be wider than the column even at an (over)generous 1.2 em per character.
                if (cell.Text.Length * cell.LargestFontSize * 1.2 <= own) continue;
                if (_anchors.ContainsKey((row, column)) || _sheet.Merges.Any(m => m.Contains(row, column))) continue;
                double width = _engine.BuildBlock(cell, cell.Text, 1e7).Width + 2 * CellPadding;
                if (width <= own + Epsilon) continue;
                var alignment = AlignmentOf(cell);
                double need = width - own;
                if (alignment == ParagraphAlignment.Right) first = Math.Min(first, Spill(row, column, -1, need));
                else if (alignment == ParagraphAlignment.Center)
                {
                    first = Math.Min(first, Spill(row, column, -1, need / 2));
                    last = Math.Max(last, Spill(row, column, 1, need / 2));
                }
                else last = Math.Max(last, Spill(row, column, 1, need + cell.Style.Indent * IndentWidth));
            }
            return range with { FirstColumn = first, LastColumn = last };
        }

        /// <summary>The last column reached when text needing <paramref name="need"/> more width spills in a direction.</summary>
        private int Spill(int row, int column, int direction, double need)
        {
            int c = column;
            // Spilled text is cut at the page edge, so there is no point reaching further than one page width.
            double reach = _sheet.ColumnWidth(column), pageWidth = (_setup.PageWidth - _setup.Margins.Left - _setup.Margins.Right) / _setup.Scale;
            for (int steps = 0; steps < 200 && need > Epsilon && reach < pageWidth; steps++)
            {
                int next = c + direction;
                if (next < 0 || next >= MaxColumns) break;
                if (_sheet.Cells.TryGetValue((row, next), out var neighbor) && neighbor.Text.Length > 0) break;
                if (_sheet.Merges.Any(m => m.Contains(row, next))) break;
                if (reach + _sheet.ColumnWidth(next) > pageWidth + Epsilon) break;
                c = next;
                need -= _sheet.ColumnWidth(c);
                reach += _sheet.ColumnWidth(c);
            }
            return c;
        }

        // ---- pictures

        private double ColumnWidthOf(int column) => _sheet.ColumnWidth(column);

        private double RowHeightOf(int row) =>
            _sheet.HiddenRows.Contains(row) ? 0 : _rowHeights.TryGetValue(row, out var h) ? h : _sheet.RowHeight(row);

        /// <summary>The distance from the corner of A1 to the start of an index, for sizes that differ from the default only at a few indexes.</summary>
        private static double StartOf(int index, double defaultSize, Func<int, double> size, IEnumerable<int> overridden)
        {
            double total = index * defaultSize;
            foreach (int key in overridden.Distinct())
                if (key < index) total += size(key) - defaultSize;
            return total;
        }

        private double ColumnStart(int column) =>
            StartOf(column, _sheet.DefaultColumnWidth, ColumnWidthOf, _sheet.ColumnWidths.Keys.Concat(_sheet.HiddenColumns));

        private double RowStart(int row) =>
            StartOf(row, _sheet.DefaultRowHeight, RowHeightOf, _sheet.RowHeights.Keys.Concat(_sheet.HiddenRows).Concat(_rowHeights.Keys));

        /// <summary>The column containing a distance from the left edge of A1.</summary>
        private int ColumnAt(double x)
        {
            double edge = 0;
            for (int c = 0; c < MaxColumns; c++)
            {
                edge += ColumnWidthOf(c);
                if (x < edge - Epsilon || c == MaxColumns - 1) return c;
            }
            return MaxColumns - 1;
        }

        /// <summary>The row containing a distance from the top edge of A1.</summary>
        private int RowAt(double y)
        {
            double edge = 0;
            for (int r = 0; r < 1_048_576; r++)
            {
                edge += RowHeightOf(r);
                if (y < edge - Epsilon || r == 1_048_575) return r;
            }
            return 1_048_575;
        }

        private IEnumerable<SheetObject> Objects() => _sheet.Images.Cast<SheetObject>().Concat(_sheet.Shapes).Concat(_sheet.Charts);

        /// <summary>Pictures have no area when either side is empty; a line may be as thin as nothing in one direction.</summary>
        private static bool IsEmpty(SheetObject obj, double width, double height) =>
            obj is SheetShape { Geometry: ShapeGeometry.Line } ? width <= 0.01 && height <= 0.01 : width <= 0.01 || height <= 0.01;

        /// <summary>The rectangle of an object measured from the corner of A1. Rotated objects are placed by their unrotated frame.</summary>
        private (double X, double Y, double Width, double Height) ImageRectFromA1(SheetObject image)
        {
            var (x, y, width, height) = image.Anchor switch
            {
                ImageAnchorKind.TwoCell => TwoCellRect(image),
                ImageAnchorKind.OneCell => (ColumnStart(image.From.Column) + image.From.ColumnOffset, RowStart(image.From.Row) + image.From.RowOffset, image.Width, image.Height),
                _ => (image.X, image.Y, image.Width, image.Height),
            };
            return PlaceInFrame(image, x, y, width, height);
        }

        private (double X, double Y, double Width, double Height) TwoCellRect(SheetObject image)
        {
            double x0 = ColumnStart(image.From.Column) + image.From.ColumnOffset, y0 = RowStart(image.From.Row) + image.From.RowOffset;
            double x1 = ColumnStart(image.To.Column) + image.To.ColumnOffset, y1 = RowStart(image.To.Row) + image.To.RowOffset;
            return (x0, y0, Math.Max(x1 - x0, 0), Math.Max(y1 - y0, 0));
        }

        /// <summary>A run of page positions that shows a contiguous stretch of the sheet: sheet distances S0–S1 are drawn from page distance P0.</summary>
        private readonly record struct Segment(double S0, double S1, double P0);

        private static List<Segment> Segments(List<int> indexes, Func<int, double> start, Func<int, double> size, Dictionary<int, double> position)
        {
            var segments = new List<Segment>();
            foreach (int index in indexes)
            {
                double s0 = start(index), s1 = s0 + size(index);
                if (segments.Count > 0 && Math.Abs(segments[^1].S1 - s0) < Epsilon)
                    segments[^1] = segments[^1] with { S1 = s1 };
                else
                    segments.Add(new Segment(s0, s1, position[index]));
            }
            return segments;
        }

        /// <summary>
        /// Draws the pictures that fall on this page. A picture is cut into the pieces that lie over the page's column and row runs
        /// (the repeated title rows and columns are runs of their own) and each piece is clipped to its run, so pictures that span
        /// several pages continue on the next one and nothing reaches outside the print range.
        /// </summary>
        private void EmitImages(List<int> columns, List<int> rows, Dictionary<int, double> columnX, Dictionary<int, double> rowY, List<LayoutItem> items)
        {
            if (_sheet.Images.Count == 0 && _sheet.Shapes.Count == 0 && _sheet.Charts.Count == 0) return;
            double originX = ColumnStart(_range.FirstColumn), originY = RowStart(_range.FirstRow);
            var columnSegments = Segments(columns, c => ColumnStart(c) - originX, ColumnWidthOf, columnX);
            var rowSegments = Segments(rows, r => RowStart(r) - originY, RowHeightOf, rowY);

            foreach (var image in Objects().OrderBy(o => o.ZOrder))
            {
                var (ax, ay, w, h) = ImageRectFromA1(image);
                double x = ax - originX, y = ay - originY;
                if (IsEmpty(image, w, h)) continue;
                // The turned bounding box of a rotated object reaches beyond its frame by at most this much on each side, and outlines and
                // text that overflow the frame reach beyond it as well.
                double reach = image.Rotation == 0 ? 0 : (Math.Abs(w) + Math.Abs(h));
                if (image is SheetShape shape)
                {
                    reach = Math.Max(reach, shape.LineWidth + 1);
                    if (shape.Glow is { } glow) reach = Math.Max(reach, glow.Radius + 1);
                    if (shape.Reflection is { } reflection) reach = Math.Max(reach, reflection.Distance + Math.Abs(h) * reflection.EndPosition + 1);
                    if (shape.Shadow is { } shade) reach = Math.Max(reach, Math.Max(Math.Abs(shade.Dx), Math.Abs(shade.Dy)) + shade.Blur + 1);
                }
                foreach (var columnSegment in columnSegments)
                {
                    double x0 = Math.Max(x - reach, columnSegment.S0), x1 = Math.Min(x + w + reach, columnSegment.S1);
                    if (x1 - x0 <= Epsilon) continue;
                    foreach (var rowSegment in rowSegments)
                    {
                        double y0 = Math.Max(y - reach, rowSegment.S0), y1 = Math.Min(y + h + reach, rowSegment.S1);
                        if (y1 - y0 <= Epsilon) continue;
                        double px = columnSegment.P0 + (x - columnSegment.S0), py = rowSegment.P0 + (y - rowSegment.S0);
                        // A turned picture sticks out of its frame, so it is only cut at the edges of the printed run, not at its frame.
                        var visible = image.Rotation != 0
                            ? (columnSegment.P0, rowSegment.P0, columnSegment.S1 - columnSegment.S0, rowSegment.S1 - rowSegment.S0)
                            : (columnSegment.P0 + (x0 - columnSegment.S0), rowSegment.P0 + (y0 - rowSegment.S0), x1 - x0, y1 - y0);
                        EmitObject(image, px, py, w, h, visible, items);
                    }
                }
            }
        }

        private void EmitObject(SheetObject obj, double x, double y, double width, double height,
            (double X, double Y, double Width, double Height) visible, List<LayoutItem> items)
        {
            if (obj is SheetShape shape) _engine.EmitShape(shape, x, y, width, height, visible, items);
            else if (obj is SheetChart chart) _engine.EmitChart(chart, x, y, width, height, visible, items);
            else if (obj is SheetImage image) EmitImage(image, x, y, width, height, visible, items);
        }

        internal static void EmitImage(SheetImage image, double x, double y, double width, double height,
            (double X, double Y, double Width, double Height) visible, List<LayoutItem> items)
        {
            int pushed = 1;
            items.Add(new PushItem(visible));
            if (image.Rotation != 0 || image.FlipH || image.FlipV)
            {
                items.Add(new PushItem(null, image.Rotation, x + width / 2, y + height / 2, image.FlipH, image.FlipV));
                pushed++;
            }
            double keepX = 1 - image.CropLeft - image.CropRight, keepY = 1 - image.CropTop - image.CropBottom;
            if (keepX > 0.001 && keepY > 0.001)
            {
                if (keepX < 1 || keepY < 1)
                {
                    // Cropped: draw the whole image so that the kept part fills the frame, and clip to the frame.
                    items.Add(new PushItem((x, y, width, height)));
                    pushed++;
                    double fullWidth = width / keepX, fullHeight = height / keepY;
                    items.Add(new ImageItem(x - image.CropLeft * fullWidth, y - image.CropTop * fullHeight, fullWidth, fullHeight, image.Image));
                }
                else items.Add(new ImageItem(x, y, width, height, image.Image));
            }
            for (int i = 0; i < pushed; i++) items.Add(new PopItem());
        }

        // ---- row heights

        /// <summary>Rows without an explicit height grow to fit wrapped or large text.</summary>
        private void ComputeRowHeights() => ComputeRowHeights(_range.FirstRow, _range.LastRow);

        private void ComputeRowHeights(int firstRow, int lastRow)
        {
            var byRow = new Dictionary<int, List<int>>();
            foreach (var (row, column) in _sheet.Cells.Keys)
            {
                if (row < firstRow || row > lastRow || column < _range.FirstColumn || column > _range.LastColumn) continue;
                if (!byRow.TryGetValue(row, out var list)) byRow[row] = list = [];
                list.Add(column);
            }
            for (int row = firstRow; row <= lastRow; row++)
            {
                if (_sheet.HiddenRows.Contains(row)) continue;
                if (_sheet.RowHeights.TryGetValue(row, out var explicitHeight))
                {
                    _rowHeights[row] = explicitHeight;
                    continue;
                }
                double height = _sheet.DefaultRowHeight;
                if (byRow.TryGetValue(row, out var columns))
                {
                    foreach (int column in columns)
                    {
                        if (_anchors.ContainsKey((row, column)) || _sheet.ColumnWidth(column) <= 0) continue;
                        var cell = _sheet.Cells[(row, column)];
                        if (cell.Text.Length == 0) continue;
                        if (IsRotated(cell.Style))
                        {
                            height = Math.Max(height, _engine.RotatedBlock(cell, _sheet.ColumnWidth(column), null).Height);
                            continue;
                        }
                        double pitch = cell.LargestFontSize * LineHeightRatio;
                        int lines = 1;
                        if (cell.Style.Wrap && cell.Kind == CellValueKind.Text)
                        {
                            var block = _engine.BuildBlock(cell, cell.Text, TextWidth(cell, _sheet.ColumnWidth(column)));
                            _blocks[(row, column)] = block;
                            height = Math.Max(height, block.Height);
                            continue;
                        }
                        height = Math.Max(height, lines * pitch);
                    }
                }
                _rowHeights[row] = Math.Min(height, MaxRowHeight);
            }
        }

        private static double TextWidth(Cell cell, double cellWidth)
        {
            double indent = cell.Style.Indent * IndentWidth;
            return Math.Max(cellWidth - 2 * CellPadding - indent - IconShift(cell), 1);
        }

        /// <summary>The icon is as large as the cell's text; the text moves right by its width plus a small gap.</summary>
        private static double IconSize(Cell cell) => cell.LargestFontSize;

        private static double IconShift(Cell cell) => cell.Icon is null ? 0 : IconSize(cell) + CellPadding;

        // ---- bands

        /// <summary>
        /// Splits the visible indexes into runs whose total size fits <paramref name="available"/>. A manual break after
        /// index b starts a new run at the next visible index; every run after the first starts with the title indexes before it.
        /// </summary>
        private static List<List<int>> Bands(List<int> indexes, Func<int, double> size, double available, SortedSet<int> breaks, List<int> titles)
        {
            double PrependSize(int first) => titles.Where(t => t < first).Sum(size);

            var bands = new List<List<int>>();
            var current = new List<int>();
            double used = 0;
            int previous = -1;
            foreach (int index in indexes)
            {
                double s = size(index);
                bool manual = current.Count > 0 && breaks.Count > 0 && breaks.GetViewBetween(previous, index - 1).Count > 0;
                if (current.Count > 0 && (manual || used + s > available + Epsilon))
                {
                    bands.Add(current);
                    current = [];
                }
                if (current.Count == 0) used = PrependSize(index);
                current.Add(index);
                used += s;
                previous = index;
            }
            if (current.Count > 0) bands.Add(current);

            // Titles are repeated in front of every run they precede.
            for (int i = 0; i < bands.Count; i++)
            {
                int first = bands[i][0];
                var repeated = titles.Where(t => t < first).ToList();
                if (repeated.Count > 0) bands[i] = [.. repeated, .. bands[i]];
            }
            return bands;
        }

        private IEnumerable<(List<int> Columns, List<int> Rows)> PageOrder(List<List<int>> columnBands, List<List<int>> rowBands)
        {
            if (_setup.PageOrderOverThenDown)
            {
                foreach (var rows in rowBands)
                    foreach (var columns in columnBands)
                        yield return (columns, rows);
            }
            else
            {
                foreach (var columns in columnBands)
                    foreach (var rows in rowBands)
                        yield return (columns, rows);
            }
        }

        // ---- drawing

        private void EmitPage(List<int> columns, List<int> rows, double scale)
        {
            var page = new LayoutPage(_setup.PageWidth, _setup.PageHeight);
            var items = new List<LayoutItem>();

            var columnX = new Dictionary<int, double>();
            double x = 0;
            foreach (int c in columns)
            {
                columnX[c] = x;
                x += _sheet.ColumnWidth(c);
            }
            double bandWidth = x;
            var rowY = new Dictionary<int, double>();
            double y = 0;
            foreach (int r in rows)
            {
                rowY[r] = y;
                y += RowHeight(r);
            }
            double bandHeight = y;

            if (_setup.PrintGridLines) EmitGridLines(columns, rows, columnX, rowY, bandWidth, bandHeight, items);

            var cells = _sheet.Cells
                .Where(kv => columnX.ContainsKey(kv.Key.Column) && rowY.ContainsKey(kv.Key.Row))
                .OrderBy(kv => kv.Key.Row).ThenBy(kv => kv.Key.Column)
                .ToList();

            var drawn = new List<(Cell Cell, int Row, int Column, double X, double Y, double Width, double Height, CellRange? Merge)>();
            foreach (var ((row, column), cell) in cells)
            {
                CellRange? merge = null;
                if (_anchors.TryGetValue((row, column), out var anchored)) merge = anchored;
                else if (_sheet.Merges.Count > 0 && _sheet.Merges.Any(m => m.Contains(row, column))) continue;

                double width, height;
                if (merge is { } m)
                {
                    width = columns.Where(c => c >= m.FirstColumn && c <= m.LastColumn).Sum(c => _sheet.ColumnWidth(c));
                    height = rows.Where(r => r >= m.FirstRow && r <= m.LastRow).Sum(RowHeight);
                }
                else
                {
                    width = _sheet.ColumnWidth(column);
                    height = RowHeight(row);
                }
                drawn.Add((cell, row, column, columnX[column], rowY[row], width, height, merge));
            }

            foreach (var d in drawn)
            {
                if (d.Cell.Style.Gradient is { } gradient) EmitGradient(gradient, d.X, d.Y, d.Width, d.Height, items);
                else if (d.Cell.Style.Fill is { } fill) items.Add(new RectItem(d.X, d.Y, d.Width, d.Height, fill));
            }
            foreach (var d in drawn)
            {
                if (d.Cell.DataBar is { } bar) EmitDataBar(bar, d.X, d.Y, d.Width, d.Height, items);
                if (d.Cell.Icon is { } icon) EmitIcon(icon, d.X + CellPadding, d.Y + (d.Height - IconSize(d.Cell)) / 2, IconSize(d.Cell), items);
            }
            foreach (var d in drawn)
                EmitText(d.Cell, d.Row, d.Column, d.X, d.Y, d.Width, d.Height, d.Merge is not null, columns, items);
            foreach (var d in drawn)
                EmitBorders(d.Cell.Style.Borders, d.X, d.Y, d.Width, d.Height, items);
            EmitImages(columns, rows, columnX, rowY, items);

            double offsetX = _setup.Margins.Left, offsetY = _setup.Margins.Top;
            if (_setup.CenterHorizontally)
                offsetX += Math.Max(0, (_setup.PageWidth - _setup.Margins.Left - _setup.Margins.Right - bandWidth * scale) / 2);
            if (_setup.CenterVertically)
                offsetY += Math.Max(0, (_setup.PageHeight - _setup.Margins.Top - _setup.Margins.Bottom - bandHeight * scale) / 2);
            foreach (var item in items) page.Items.Add(Transform(item, offsetX, offsetY, scale));
            _document.Pages.Add(page);
            _records.Add(new PageRecord(page, _sheet, _sheetPages++, scale));
        }

        private static readonly Color White = new(255, 255, 255);

        private static void EmitDataBar(CellDataBar bar, double x, double y, double width, double height, List<LayoutItem> items)
        {
            double left = x + bar.Start * width, barWidth = (bar.End - bar.Start) * width;
            if (barWidth <= 0.05 || height <= 2) return;
            double top = y + 1, barHeight = height - 2;
            if (!bar.Gradient) items.Add(new RectItem(left, top, barWidth, barHeight, bar.Color));
            else
            {
                // Gradient bars are solid at their base (next to the axis) and fade towards the free end.
                var light = bar.Color.Blend(White, 0.8);
                items.Add(new GradientItem(left, top, barWidth, barHeight, bar.Negative != bar.Mirrored ? 180 : 0, [new GradientStop(0, bar.Color), new GradientStop(1, light)]));
            }
            if (bar.Axis is { } axis)
                items.Add(new RectItem(x + axis * width - 0.375, top, 0.75, barHeight, bar.AxisColor ?? Color.Black));
            if (bar.BorderColor is { } border)
            {
                // A thin solid outline, drawn inside the bar.
                double line = Math.Min(0.75, Math.Min(barWidth, barHeight) / 2);
                items.Add(new RectItem(left, top, barWidth, line, border));
                items.Add(new RectItem(left, top + barHeight - line, barWidth, line, border));
                items.Add(new RectItem(left, top + line, line, barHeight - 2 * line, border));
                items.Add(new RectItem(left + barWidth - line, top + line, line, barHeight - 2 * line, border));
            }
        }

        /// <summary>Draws an icon in a square of the given size.</summary>
        private static void EmitIcon(CellIcon icon, double x, double y, double size, List<LayoutItem> items)
        {
            if (size <= 0) return;
            LayoutItem Poly(Color color, params (double X, double Y)[] points) =>
                new PolygonItem(points.Select(p => (x + p.X * size, y + p.Y * size)).ToList(), color);
            LayoutItem Disc(Color color, double l, double t, double w, double h) => new EllipseItem(x + l * size, y + t * size, w * size, h * size, color);
            // Polygon turned clockwise by the given degrees about the center of the icon square.
            LayoutItem Turned(Color color, double degrees, params (double X, double Y)[] points)
            {
                double radians = degrees * Math.PI / 180, cos = Math.Cos(radians), sin = Math.Sin(radians);
                return Poly(color, points.Select(p => (0.5 + (p.X - 0.5) * cos - (p.Y - 0.5) * sin, 0.5 + (p.X - 0.5) * sin + (p.Y - 0.5) * cos)).ToArray());
            }
            (double X, double Y)[] Star() =>
                Enumerable.Range(0, 10).Select(k =>
                {
                    double angle = (k * 36 - 90) * Math.PI / 180, radius = k % 2 == 0 ? 0.5 : 0.21;
                    return (0.5 + radius * Math.Cos(angle), 0.55 + radius * Math.Sin(angle));
                }).ToArray();
            var arrowUp = new (double X, double Y)[] { (0.5, 0), (1, 0.5), (0.72, 0.5), (0.72, 1), (0.28, 1), (0.28, 0.5), (0, 0.5) };
            var arrowDown = new (double X, double Y)[] { (0.5, 1), (1, 0.5), (0.72, 0.5), (0.72, 0), (0.28, 0), (0.28, 0.5), (0, 0.5) };
            var check = new (double X, double Y)[] { (0.2, 0.52), (0.31, 0.41), (0.43, 0.54), (0.69, 0.27), (0.8, 0.38), (0.43, 0.76) };
            var bang = new (double X, double Y)[] { (0.43, 0.2), (0.57, 0.2), (0.55, 0.58), (0.45, 0.58) };
            var cross = new (double X, double Y)[] { (0.28, 0.36), (0.36, 0.28), (0.5, 0.42), (0.64, 0.28), (0.72, 0.36), (0.58, 0.5), (0.72, 0.64), (0.64, 0.72), (0.5, 0.58), (0.36, 0.72), (0.28, 0.64), (0.42, 0.5) };
            var empty = new Color(0xD9, 0xD9, 0xD9);
            switch (icon.Shape)
            {
                case IconShape.ArrowUp:
                    items.Add(Poly(icon.Color, arrowUp));
                    break;
                case IconShape.ArrowDown:
                    items.Add(Poly(icon.Color, arrowDown));
                    break;
                case IconShape.ArrowSideways:
                    items.Add(Poly(icon.Color, (1, 0.5), (0.5, 0), (0.5, 0.28), (0, 0.28), (0, 0.72), (0.5, 0.72), (0.5, 1)));
                    break;
                case IconShape.ArrowUpRight:
                    items.Add(Turned(icon.Color, 45, arrowUp));
                    break;
                case IconShape.ArrowDownRight:
                    items.Add(Turned(icon.Color, -45, arrowDown));
                    break;
                case IconShape.Circle:
                    items.Add(Disc(icon.Color, 0, 0, 1, 1));
                    break;
                case IconShape.RimmedCircle:
                    items.Add(Disc(Color.Black, 0, 0, 1, 1));
                    items.Add(Disc(icon.Color, 0.12, 0.12, 0.76, 0.76));
                    break;
                case IconShape.Check:
                    items.Add(Disc(icon.Color, 0, 0, 1, 1));
                    items.Add(Poly(White, check));
                    break;
                case IconShape.Exclamation:
                    items.Add(Disc(icon.Color, 0, 0, 1, 1));
                    items.Add(Poly(White, bang));
                    items.Add(Disc(White, 0.42, 0.65, 0.16, 0.16));
                    break;
                case IconShape.Cross:
                    items.Add(Disc(icon.Color, 0, 0, 1, 1));
                    items.Add(Poly(White, cross));
                    break;
                case IconShape.BareCheck:
                    items.Add(Poly(icon.Color, (0, 0.52), (0.15, 0.37), (0.36, 0.58), (0.84, 0.08), (1, 0.23), (0.36, 0.9)));
                    break;
                case IconShape.BareExclamation:
                    items.Add(Poly(icon.Color, (0.38, 0), (0.62, 0), (0.58, 0.64), (0.42, 0.64)));
                    items.Add(Disc(icon.Color, 0.37, 0.74, 0.26, 0.26));
                    break;
                case IconShape.BareCross:
                    items.Add(Poly(icon.Color, (0.05, 0.2), (0.2, 0.05), (0.5, 0.35), (0.8, 0.05), (0.95, 0.2), (0.65, 0.5), (0.95, 0.8), (0.8, 0.95), (0.5, 0.65), (0.2, 0.95), (0.05, 0.8), (0.35, 0.5)));
                    break;
                case IconShape.Flag:
                    items.Add(Poly(icon.Color, (0.12, 0), (0.24, 0), (0.24, 1), (0.12, 1)));
                    items.Add(Poly(icon.Color, (0.24, 0.05), (0.95, 0.28), (0.24, 0.55)));
                    break;
                case IconShape.Diamond:
                    items.Add(Poly(icon.Color, (0.5, 0), (1, 0.5), (0.5, 1), (0, 0.5)));
                    break;
                case IconShape.TriangleUp:
                    items.Add(Poly(icon.Color, (0.5, 0.08), (1, 0.92), (0, 0.92)));
                    break;
                case IconShape.TriangleDown:
                    items.Add(Poly(icon.Color, (0, 0.08), (1, 0.08), (0.5, 0.92)));
                    break;
                case IconShape.Dash:
                    items.Add(Poly(icon.Color, (0, 0.35), (1, 0.35), (1, 0.65), (0, 0.65)));
                    break;
                case IconShape.Star:
                    items.Add(Poly(empty, Star()));
                    if (icon.Level >= 2) items.Add(Poly(icon.Color, Star()));
                    else if (icon.Level == 1)
                    {
                        items.Add(new PushItem((x, y, size / 2, size)));
                        items.Add(Poly(icon.Color, Star()));
                        items.Add(new PopItem());
                    }
                    break;
                case IconShape.Quarters:
                    items.Add(Disc(empty, 0, 0, 1, 1));
                    for (int q = 0; q < Math.Clamp(icon.Level, 0, 4); q++)
                    {
                        var wedge = new List<(double X, double Y)> { (0.5, 0.5) };
                        for (int step = 0; step <= 6; step++)
                        {
                            double angle = (q * 90 + step * 15 - 90) * Math.PI / 180;
                            wedge.Add((0.5 + 0.5 * Math.Cos(angle), 0.5 + 0.5 * Math.Sin(angle)));
                        }
                        items.Add(Poly(icon.Color, wedge.ToArray()));
                    }
                    break;
                case IconShape.Bars:
                    for (int bar = 0; bar < 4; bar++)
                    {
                        double height = 0.25 * (bar + 1), left = 0.04 + bar * 0.25;
                        items.Add(new RectItem(x + left * size, y + (1 - height) * size, 0.18 * size, height * size, bar < icon.Level ? icon.Color : empty));
                    }
                    break;
                case IconShape.Boxes:
                {
                    items.Add(new RectItem(x + 0.05 * size, y + 0.05 * size, 0.9 * size, 0.9 * size, empty));
                    (double L, double T)[] quadrants = [(0.05, 0.05), (0.5, 0.05), (0.5, 0.5), (0.05, 0.5)];
                    for (int q = 0; q < Math.Clamp(icon.Level, 0, 4); q++)
                        items.Add(new RectItem(x + quadrants[q].L * size, y + quadrants[q].T * size, 0.45 * size, 0.45 * size, icon.Color));
                    break;
                }
            }
        }

        /// <summary>
        /// Linear gradients become a shaded rectangle. Excel's path gradients are rectangular: the color depends on how far a point
        /// is from the focus relative to the farthest edge, measured per axis, so they are drawn as nested rectangles.
        /// </summary>
        internal static void EmitGradient(GradientFill gradient, double x, double y, double width, double height, List<LayoutItem> items)
        {
            if (width <= 0 || height <= 0) return;
            if (gradient.Kind == GradientKind.Linear)
            {
                items.Add(new GradientItem(x, y, width, height, gradient.Degree, gradient.Stops));
                return;
            }
            double fx = x + width * (gradient.Left + gradient.Right) / 2, fy = y + height * (gradient.Top + gradient.Bottom) / 2;
            double reachX = Math.Max(Math.Max(fx - x, x + width - fx), 1e-6), reachY = Math.Max(Math.Max(fy - y, y + height - fy), 1e-6);
            int steps = Math.Clamp((int)(Math.Max(width, height) / 2), 16, 64);
            items.Add(new PushItem((x, y, width, height)));
            for (int k = steps - 1; k >= 0; k--)
            {
                double outer = (k + 1.0) / steps;
                items.Add(new RectItem(fx - outer * reachX, fy - outer * reachY, 2 * outer * reachX, 2 * outer * reachY, gradient.ColorAt((k + 0.5) / steps)));
            }
            items.Add(new PopItem());
        }

        private void EmitGridLines(List<int> columns, List<int> rows, Dictionary<int, double> columnX, Dictionary<int, double> rowY,
            double width, double height, List<LayoutItem> items)
        {
            foreach (int c in columns)
                items.Add(new LineItem(columnX[c], 0, columnX[c], height, 0.5, GridColor));
            items.Add(new LineItem(width, 0, width, height, 0.5, GridColor));
            foreach (int r in rows)
                items.Add(new LineItem(0, rowY[r], width, rowY[r], 0.5, GridColor));
            items.Add(new LineItem(0, height, width, height, 0.5, GridColor));
        }

        private static void EmitBorders(BorderSet borders, double x, double y, double w, double h, List<LayoutItem> items)
        {
            if (!borders.HasAny) return;
            Edge(borders.Top, x, y, x + w, y, items);
            Edge(borders.Bottom, x, y + h, x + w, y + h, items);
            Edge(borders.Left, x, y, x, y + h, items);
            Edge(borders.Right, x + w, y, x + w, y + h, items);
        }

        private static void Edge(Border border, double x1, double y1, double x2, double y2, List<LayoutItem> items)
        {
            if (!border.IsVisible) return;
            var color = border.Color ?? Color.Black;
            var style = border.Style switch { BorderStyle.Dotted => LineStyle.Dotted, BorderStyle.Dashed => LineStyle.Dashed, _ => LineStyle.Solid };
            if (border.Style == BorderStyle.Double)
            {
                double line = border.Width / 3;
                bool horizontal = y1 == y2;
                double d = line; // lines sit one line width either side of the edge
                items.Add(horizontal ? new LineItem(x1, y1 - d, x2, y2 - d, line, color) : new LineItem(x1 - d, y1, x2 - d, y2, line, color));
                items.Add(horizontal ? new LineItem(x1, y1 + d, x2, y2 + d, line, color) : new LineItem(x1 + d, y1, x2 + d, y2, line, color));
                return;
            }
            items.Add(new LineItem(x1, y1, x2, y2, border.Width, color, style));
        }

        /// <summary>The width of empty neighbors (in the given direction) that text may spill into, up to <paramref name="need"/>.</summary>
        private double Room(int row, List<int> columns, int position, int direction, double need)
        {
            double room = 0;
            for (int i = position + direction; i >= 0 && i < columns.Count && room < need - Epsilon; i += direction)
            {
                int c = columns[i];
                if (_sheet.Cells.TryGetValue((row, c), out var neighbor) && neighbor.Text.Length > 0) break;
                if (_sheet.Merges.Any(m => m.Contains(row, c))) break;
                room += _sheet.ColumnWidth(c);
            }
            return room;
        }

        /// <summary>
        /// Draws rotated or stacked text: the unrotated block is positioned so that its rotated bounding box sits in the cell
        /// according to the cell's alignment (Excel aligns the box, not the text), turned about the box center and clipped to the cell.
        /// </summary>
        private void EmitRotated(Cell cell, double x, double y, double width, double height, List<LayoutItem> items)
        {
            var style = cell.Style;
            var (block, boxWidth, boxHeight) = _engine.RotatedBlock(cell, width, height);
            double textWidth = block.Width, textHeight = block.Height;
            var alignment = AlignmentOf(cell);
            double indent = style.Indent * IndentWidth;
            double left = alignment switch
            {
                ParagraphAlignment.Right => x + width - CellPadding - boxWidth,
                ParagraphAlignment.Center => x + (width - boxWidth) / 2,
                _ => x + CellPadding + indent,
            };
            double top = style.VerticalAlignment switch
            {
                CellVerticalAlignment.Top => y,
                CellVerticalAlignment.Center => y + (height - boxHeight) / 2,
                _ => y + height - boxHeight,
            };
            double cx = left + boxWidth / 2, cy = top + boxHeight / 2;
            items.Add(new PushItem((x, y, width, height), style.Stacked ? 0 : style.Rotation, cx, cy));
            double lineTop = cy - textHeight / 2, blockLeft = cx - textWidth / 2;
            for (int i = 0; i < block.Lines.Count; i++)
            {
                var (pitch, ascent, descent) = block.Metrics[i];
                var line = block.Lines[i];
                if (line.Atoms.Count > 0)
                {
                    double baseline = lineTop + (pitch - (ascent + descent)) / 2 + ascent;
                    double x0 = style.Stacked ? blockLeft + (textWidth - LineWidth(line)) / 2 : blockLeft;
                    _engine.EmitLine(line, x0, baseline, double.NegativeInfinity, double.PositiveInfinity, items);
                }
                lineTop += pitch;
            }
            items.Add(new PopItem());
        }

        private void EmitText(Cell cell, int row, int column, double x, double y, double width, double height, bool merged, List<int> columns, List<LayoutItem> items)
        {
            if (cell.Text.Length == 0) return;
            var style = cell.Style;
            if (IsRotated(style))
            {
                EmitRotated(cell, x, y, width, height, items);
                return;
            }
            bool wrap = style.Wrap && cell.Kind == CellValueKind.Text;
            double textWidth = TextWidth(cell, width);

            string text = cell.Text;
            TextBlock block;
            if (!merged && _blocks.TryGetValue((row, column), out var cached)) block = cached;
            else block = _engine.BuildBlock(cell, text, textWidth);

            if (cell.Kind == CellValueKind.Number && block.Width > textWidth + Epsilon)
            {
                // Numbers never spill: show as many '#' as fit.
                var hash = _engine._fonts.ForCharacter(style.Font, '#');
                double hashWidth = Math.Max(Itemizer.Measure(hash, "#", style.Font.FontSize, style.Font, false), 0.1);
                text = new string('#', Math.Max(1, (int)(textWidth / hashWidth)));
                block = _engine.BuildBlock(cell, text, textWidth);
            }

            if (style.HorizontalAlignment == HorizontalAlignment.Fill && !wrap && block.Width > 0.01 && block.Width <= textWidth + Epsilon && text.Length > 0)
            {
                // Fill alignment repeats the text as often as it fits the cell.
                int copies = (int)Math.Floor((textWidth + Epsilon) / block.Width);
                while (copies > 1)
                {
                    var repeated = _engine.BuildBlock(cell, string.Concat(Enumerable.Repeat(text, copies)), textWidth);
                    if (repeated.Width <= textWidth + Epsilon) { block = repeated; break; }
                    copies--;
                }
            }

            var alignment = AlignmentOf(cell);
            double indent = style.Indent * IndentWidth;
            double left = x + CellPadding + IconShift(cell) + (alignment == ParagraphAlignment.Left || alignment == ParagraphAlignment.Justify ? indent : 0);
            double right = x + width - CellPadding - (alignment == ParagraphAlignment.Right ? indent : 0);

            double clipLeft = x, clipRight = x + width;
            if (!wrap && !merged && cell.Kind == CellValueKind.Text && block.Width + 2 * CellPadding > width + Epsilon)
            {
                double need = block.Width + 2 * CellPadding + (alignment == ParagraphAlignment.Left ? indent : 0) - width;
                int position = columns.IndexOf(column);
                if (alignment == ParagraphAlignment.Right) clipLeft -= Room(row, columns, position, -1, need);
                else if (alignment == ParagraphAlignment.Center)
                {
                    double half = Math.Min(Room(row, columns, position, -1, need / 2), Room(row, columns, position, 1, need / 2));
                    clipLeft -= half;
                    clipRight += half;
                }
                else clipRight += Room(row, columns, position, 1, need);
            }

            double blockHeight = block.Height;
            double top = style.VerticalAlignment switch
            {
                CellVerticalAlignment.Top => y,
                CellVerticalAlignment.Center => y + (height - blockHeight) / 2,
                _ => y + height - blockHeight,
            };
            double lineTop = top;
            for (int i = 0; i < block.Lines.Count; i++, lineTop += block.Metrics[i - 1].Pitch)
            {
                var (pitch, ascent, descent) = block.Metrics[i];
                if (lineTop < y - 0.5 || lineTop + pitch > y + height + 0.5) continue; // outside the cell vertically
                var line = block.Lines[i];
                if (line.Atoms.Count == 0) continue;
                double baseline = lineTop + (pitch - (ascent + descent)) / 2 + ascent;
                double x0;
                if (wrap) x0 = left;
                else
                {
                    double w = LineWidth(line);
                    x0 = alignment switch
                    {
                        ParagraphAlignment.Right => right - w,
                        ParagraphAlignment.Center => (left + right) / 2 - w / 2,
                        _ => left,
                    };
                }
                _engine.EmitLine(line, x0, baseline, clipLeft, clipRight, items);
            }
        }
    }
}
