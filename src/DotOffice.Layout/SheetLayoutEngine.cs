using DotOffice.Core;

namespace DotOffice.Layout;

/// <summary>
/// Lays out the worksheets of a <see cref="Workbook"/> into pages: each sheet's print range is split into bands of
/// columns and rows that fit a page, and every cell is drawn into its band with its fill, borders and text.
/// </summary>
public sealed class SheetLayoutEngine
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
        for (int i = 0; i < records.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            EmitHeaderFooter(records[i], i + 1, records.Count, workbook);
        }
        return document;
    }

    /// <summary>A page of a sheet, kept until headers and footers can be drawn.</summary>
    private sealed record PageRecord(LayoutPage Page, Worksheet Sheet, int IndexInSheet);

    // ------------------------------------------------------------------ headers and footers

    private static readonly RunProperties HeaderFont = new() { FontFamily = "Calibri", FontSize = 11 };

    private void EmitHeaderFooter(PageRecord record, int number, int total, Workbook workbook)
    {
        var setup = record.Sheet.PageSetup;
        bool first = record.IndexInSheet == 0;
        static HeaderFooterParts Pick(HeaderFooterText text, bool first, int number) =>
            first && text.First is { } f ? f : number % 2 == 0 && text.Even is { } e ? e : text.Odd;
        var header = Pick(setup.Header, first, number);
        var footer = Pick(setup.Footer, first, number);
        if (header.IsEmpty && footer.IsEmpty) return;
        var context = new FieldContext(number, total, record.Sheet.Name, workbook.Title ?? "");
        var page = record.Page;
        double left = setup.Margins.Left, right = page.Width - setup.Margins.Right;

        void Part(string text, double anchorLeft, double anchorRight, int alignment, bool isHeader)
        {
            if (string.IsNullOrEmpty(text)) return;
            var lines = text.Replace("\r\n", "\n").Split('\n').Select(l => BuildHeaderLine(l, context)).Where(l => l is not null).Select(l => l!).ToList();
            if (lines.Count == 0) return;
            double pitch(HeaderLine l) => l.Size * 1.2;
            double height = lines.Sum(pitch);
            double top = isHeader ? setup.HeaderDistance : page.Height - setup.FooterDistance - height;
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

        foreach (var (parts, isHeader) in new[] { (header, true), (footer, false) })
        {
            Part(parts.Left, left, right, 0, isHeader);
            Part(parts.Center, left, right, 1, isHeader);
            Part(parts.Right, left, right, 2, isHeader);
        }
    }

    private readonly record struct FieldContext(int Page, int Pages, string Sheet, string File);

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
            runs.Add((buffer.ToString(), props));
            maxSize = Math.Max(maxSize, props.FontSize);
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
        public double Pitch { get; init; }
        public double Ascent { get; init; }
        public double Descent { get; init; }
        public double Width { get; set; }
        public double Height => Lines.Count * Pitch;
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

    /// <summary>Breaks a cell's text into lines; <paramref name="wrapWidth"/> is the text width for wrapping cells.</summary>
    private TextBlock BuildBlock(Cell cell, string text, double wrapWidth)
    {
        var font = cell.Style.Font;
        bool wrap = cell.Style.Wrap && cell.Kind == CellValueKind.Text;
        var metrics = FontMetrics.Of(_fonts.ForCharacter(font, '0').Font, font.FontSize);
        var block = new TextBlock { Pitch = Math.Max(font.FontSize * LineHeightRatio, 1), Ascent = metrics.Ascent, Descent = metrics.Descent };

        IEnumerable<string> segments = wrap
            ? text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
            : [text.Replace("\r", "").Replace("\n", "")];
        var alignment = wrap ? AlignmentOf(cell) : ParagraphAlignment.Left;
        var frame = new LineFrame(wrap ? Math.Max(wrapWidth, 1) : 1e7, 0, 0, 0, [], 36);
        foreach (var segment in segments)
        {
            if (segment.Length == 0)
            {
                block.Lines.Add(new Line { Atoms = [] });
                continue;
            }
            var paragraph = new Paragraph { MarkProperties = font };
            paragraph.Inlines.Add(new TextRun(segment, font));
            var atoms = _itemizer.Itemize(paragraph, _ => "");
            var lines = LineBuilder.Build(atoms, frame, alignment, false, null);
            if (lines.Count == 0) block.Lines.Add(new Line { Atoms = [] });
            else block.Lines.AddRange(lines);
        }
        foreach (var line in block.Lines) block.Width = Math.Max(block.Width, LineWidth(line));
        return block;
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
            if (FindRange() is not { } range) return;
            _range = range;
            foreach (var merge in _sheet.Merges) _anchors[(merge.FirstRow, merge.FirstColumn)] = merge;

            var columns = Enumerable.Range(_range.FirstColumn, _range.ColumnCount).Where(c => _sheet.ColumnWidth(c) > 0).ToList();
            ComputeRowHeights();
            var rows = Enumerable.Range(_range.FirstRow, _range.RowCount).Where(r => RowHeight(r) > 0).ToList();
            if (columns.Count == 0 || rows.Count == 0) return;

            var titleColumns = _setup.TitleColumns is { } tc ? columns.Where(c => c >= tc.From && c <= tc.To).ToList() : [];
            var titleRows = _setup.TitleRows is { } tr ? rows.Where(r => r >= tr.From && r <= tr.To).ToList() : [];

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

        private CellRange? FindRange()
        {
            int r0 = int.MaxValue, c0 = int.MaxValue, r1 = -1, c1 = -1;
            foreach (var (row, column) in _sheet.Cells.Keys)
            {
                r0 = Math.Min(r0, row);
                c0 = Math.Min(c0, column);
                r1 = Math.Max(r1, row);
                c1 = Math.Max(c1, column);
            }
            foreach (var merge in _sheet.Merges)
            {
                if (r1 < 0) continue;
                r1 = Math.Max(r1, merge.LastRow);
                c1 = Math.Max(c1, merge.LastColumn);
            }
            if (_setup.PrintArea is { } area)
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
                if (cell.Kind != CellValueKind.Text || cell.Style.Wrap || cell.Text.Length == 0) continue;
                double own = _sheet.ColumnWidth(column);
                if (cell.Text.Length * cell.Style.Font.FontSize * 0.25 <= own) continue;
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

        // ---- row heights

        /// <summary>Rows without an explicit height grow to fit wrapped or large text.</summary>
        private void ComputeRowHeights()
        {
            var byRow = new Dictionary<int, List<int>>();
            foreach (var (row, column) in _sheet.Cells.Keys)
            {
                if (!_range.Contains(row, column)) continue;
                if (!byRow.TryGetValue(row, out var list)) byRow[row] = list = [];
                list.Add(column);
            }
            for (int row = _range.FirstRow; row <= _range.LastRow; row++)
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
                        double pitch = cell.Style.Font.FontSize * LineHeightRatio;
                        int lines = 1;
                        if (cell.Style.Wrap && cell.Kind == CellValueKind.Text)
                        {
                            var block = _engine.BuildBlock(cell, cell.Text, TextWidth(cell, _sheet.ColumnWidth(column)));
                            _blocks[(row, column)] = block;
                            lines = Math.Max(block.Lines.Count, 1);
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
            return Math.Max(cellWidth - 2 * CellPadding - indent, 1);
        }

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
                if (d.Cell.Style.Fill is { } fill) items.Add(new RectItem(d.X, d.Y, d.Width, d.Height, fill));
            foreach (var d in drawn)
                EmitText(d.Cell, d.Row, d.Column, d.X, d.Y, d.Width, d.Height, d.Merge is not null, columns, items);
            foreach (var d in drawn)
                EmitBorders(d.Cell.Style.Borders, d.X, d.Y, d.Width, d.Height, items);

            double offsetX = _setup.Margins.Left, offsetY = _setup.Margins.Top;
            if (_setup.CenterHorizontally)
                offsetX += Math.Max(0, (_setup.PageWidth - _setup.Margins.Left - _setup.Margins.Right - bandWidth * scale) / 2);
            if (_setup.CenterVertically)
                offsetY += Math.Max(0, (_setup.PageHeight - _setup.Margins.Top - _setup.Margins.Bottom - bandHeight * scale) / 2);
            foreach (var item in items) page.Items.Add(Transform(item, offsetX, offsetY, scale));
            _document.Pages.Add(page);
            _records.Add(new PageRecord(page, _sheet, _sheetPages++));
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

        private void EmitText(Cell cell, int row, int column, double x, double y, double width, double height, bool merged, List<int> columns, List<LayoutItem> items)
        {
            if (cell.Text.Length == 0) return;
            var style = cell.Style;
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

            var alignment = AlignmentOf(cell);
            double indent = style.Indent * IndentWidth;
            double left = x + CellPadding + (alignment == ParagraphAlignment.Left || alignment == ParagraphAlignment.Justify ? indent : 0);
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
            for (int i = 0; i < block.Lines.Count; i++)
            {
                double lineTop = top + i * block.Pitch;
                if (lineTop < y - 0.5 || lineTop + block.Pitch > y + height + 0.5) continue; // outside the cell vertically
                var line = block.Lines[i];
                if (line.Atoms.Count == 0) continue;
                double baseline = lineTop + (block.Pitch - (block.Ascent + block.Descent)) / 2 + block.Ascent;
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
