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
        foreach (var sheet in workbook.Sheets)
        {
            if (sheet.Visibility != SheetVisibility.Visible) continue;
            new SheetPass(this, workbook, sheet, document, cancellationToken).Run();
        }
        if (document.Pages.Count == 0) document.Pages.Add(new LayoutPage(612, 792));
        return document;
    }

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
        private readonly CancellationToken _ct;
        private readonly SheetPageSetup _setup;

        private CellRange _range;
        private readonly Dictionary<int, double> _rowHeights = [];
        private readonly Dictionary<(int Row, int Column), TextBlock> _blocks = [];
        private readonly Dictionary<(int Row, int Column), CellRange> _anchors = [];

        public SheetPass(SheetLayoutEngine engine, Workbook workbook, Worksheet sheet, LayoutDocument document, CancellationToken ct)
        {
            _engine = engine;
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

            double scale = _setup.Scale;
            double availWidth = (_setup.PageWidth - _setup.Margins.Left - _setup.Margins.Right) / scale;
            double availHeight = (_setup.PageHeight - _setup.Margins.Top - _setup.Margins.Bottom) / scale;
            var columnBands = Bands(columns, c => _sheet.ColumnWidth(c), availWidth);
            var rowBands = Bands(rows, RowHeight, availHeight);

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
            return r1 < 0 ? null : new CellRange(r0, c0, r1, c1);
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

        /// <summary>Splits the visible indexes into runs whose total size fits <paramref name="available"/>.</summary>
        private static List<List<int>> Bands(List<int> indexes, Func<int, double> size, double available)
        {
            var bands = new List<List<int>>();
            var current = new List<int>();
            double used = 0;
            foreach (int index in indexes)
            {
                double s = size(index);
                if (current.Count > 0 && used + s > available + Epsilon)
                {
                    bands.Add(current);
                    current = [];
                    used = 0;
                }
                current.Add(index);
                used += s;
            }
            if (current.Count > 0) bands.Add(current);
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
                EmitText(d.Cell, d.Row, d.Column, d.X, d.Y, d.Width, d.Height, d.Merge is not null, items);
            foreach (var d in drawn)
                EmitBorders(d.Cell.Style.Borders, d.X, d.Y, d.Width, d.Height, items);

            double offsetX = _setup.Margins.Left, offsetY = _setup.Margins.Top;
            if (_setup.CenterHorizontally)
                offsetX += Math.Max(0, (_setup.PageWidth - _setup.Margins.Left - _setup.Margins.Right - bandWidth * scale) / 2);
            if (_setup.CenterVertically)
                offsetY += Math.Max(0, (_setup.PageHeight - _setup.Margins.Top - _setup.Margins.Bottom - bandHeight * scale) / 2);
            foreach (var item in items) page.Items.Add(Transform(item, offsetX, offsetY, scale));
            _document.Pages.Add(page);
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

        private void EmitText(Cell cell, int row, int column, double x, double y, double width, double height, bool merged, List<LayoutItem> items)
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
