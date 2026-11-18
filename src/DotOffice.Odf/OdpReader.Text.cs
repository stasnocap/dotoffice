using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DotOffice.Core;

namespace DotOffice.Odf;

internal sealed partial class OdpContext
{
    private static readonly Regex OutlineRoot = new(@"^(.*-outline)\d$", RegexOptions.Compiled);

    private sealed class TextState
    {
        public bool LastSpace = true;
        public int Characters;
    }

    // ------------------------------------------------------------------ text

    /// <summary>The paragraphs of a text box, shape or table cell. <paramref name="g"/> holds the object's merged style properties.</summary>
    private List<ShapeParagraph> ReadText(XElement container, XElement owner, Dictionary<string, string> g, string? cls)
    {
        var result = new List<ShapeParagraph>();
        var shapeText = (string?)owner.Attribute(Ns.Draw + "text-style-name");
        var baseProps = shapeText is null ? g : Merge(g, styles.Properties("paragraph", shapeText, false));
        var counters = new int[12];
        var embedded = cls == "outline" ? EmbeddedListStyle(owner) : null;
        Walk(container, 0, null, embedded);
        // A text of empty paragraphs only carries no text.
        if (result.All(p => p.Runs.Count == 0 || p.Runs.All(r => r.Text.Length == 0))) result.Clear();
        return result;

        void Walk(XElement parent, int level, XElement? listStyle, XElement? fallbackList)
        {
            if (level > 10) return;
            foreach (var child in parent.Elements())
            {
                if (result.Count > 20_000) return;
                if (child.Name.Namespace != Ns.Text) continue;
                switch (child.Name.LocalName)
                {
                    case "p" or "h": Paragraph(child, level, listStyle ?? fallbackList); break;
                    case "list":
                    {
                        var style = styles.ListStyle((string?)child.Attribute(Ns.Text + "style-name")) ?? listStyle;
                        int next = level + 1;
                        if ((string?)child.Attribute(Ns.Text + "continue-numbering") != "true" && next < counters.Length) counters[next] = 0;
                        foreach (var item in child.Elements().Where(i => i.Name == Ns.Text + "list-item" || i.Name == Ns.Text + "list-header"))
                        {
                            if (next < counters.Length)
                            {
                                counters[next] = int.TryParse((string?)item.Attribute(Ns.Text + "start-value"), out var sv) ? sv : counters[next] + 1;
                                for (int l = next + 1; l < counters.Length; l++) counters[l] = 0;
                            }
                            Walk(item, next, style, fallbackList);
                        }
                        break;
                    }
                }
            }
        }

        void Paragraph(XElement p, int level, XElement? listStyle)
        {
            var levelBase = cls == "outline" && level > 1 ? LevelProps(owner, level, baseProps) : baseProps;
            var props = Merge(levelBase, styles.Properties("paragraph", (string?)p.Attribute(Ns.Text + "style-name"), false));
            var run = OdfText.Run(styles, props, new RunProperties { FontFamily = "Liberation Serif", FontSize = 18 });
            var runs = new List<CellTextRun>();
            Inline(p, run, null, runs, new TextState());
            double size = runs.Count > 0 ? runs.Max(r => r.Font.FontSize) : run.FontSize;

            string? P(string key) => props.GetValueOrDefault("paragraph-properties/" + key);
            double ml = OdfUnits.Length(P("margin-left")) ?? 0, ti = OdfUnits.Length(P("text-indent")) ?? 0;
            double indent = ml + ti, margin = ml;
            if (level > 0 && ListLevel(listStyle, level) is { } levelEl)
            {
                var lp = levelEl.Element(Ns.Style + "list-level-properties");
                var la = lp?.Element(Ns.Style + "list-level-label-alignment");
                double left, first;
                if (la is not null)
                {
                    left = OdfUnits.Length((string?)la.Attribute(Ns.Fo + "margin-left")) ?? 0;
                    first = left + (OdfUnits.Length((string?)la.Attribute(Ns.Fo + "text-indent")) ?? 0);
                }
                else
                {
                    double before = OdfUnits.Length((string?)lp?.Attribute(Ns.Text + "space-before")) ?? 0;
                    double width = OdfUnits.Length((string?)lp?.Attribute(Ns.Text + "min-label-width")) ?? 0;
                    left = ml + before + width;
                    first = ml + before;
                }
                (indent, margin) = (first, left);
                if (Label(levelEl, listStyle!, level, counters) is { Length: > 0 } label)
                {
                    var labelRun = run with { Underline = UnderlineStyle.None, Strike = false };
                    var tp = levelEl.Element(Ns.Style + "text-properties");
                    if (tp?.Attribute(Ns.Fo + "font-size") is { Value: var fs } && fs.EndsWith('%') && label != "•"
                        && double.TryParse(fs.AsSpan(0, fs.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct) && pct > 50 && pct < 100)
                        labelRun = labelRun with { FontSize = run.FontSize * pct / 100 };
                    if (tp?.Attribute(Ns.Fo + "color") is { } c && OdfUnits.Hex(c.Value) is { } color) labelRun = labelRun with { Color = color };
                    string gap = margin > indent + 0.5 ? "\t" : " ";
                    runs.Insert(0, new CellTextRun(label + gap, labelRun));
                }
            }
            indent = Math.Max(0, indent);
            margin = Math.Max(0, margin);

            var spacing = 1.0;
            double? pitch = null;
            if (P("line-height") is { } lh)
            {
                if (lh.EndsWith('%') && double.TryParse(lh.AsSpan(0, lh.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct) && pct > 0) spacing = pct / 100;
                else if (OdfUnits.Length(lh) is { } len and > 0) pitch = len;
            }
            var align = P("text-align") switch
            {
                "center" => HorizontalAlignment.Center, "end" or "right" => HorizontalAlignment.Right, "justify" => HorizontalAlignment.Justify, _ => HorizontalAlignment.Left,
            };
            if (P("writing-mode") is "rl-tb" or "rl" && align == HorizontalAlignment.Left) align = HorizontalAlignment.Right;
            result.Add(new ShapeParagraph(align, runs, size, indent, Math.Abs(margin - indent) < 0.01 ? null : margin, spacing, pitch,
                Math.Max(0, OdfUnits.Length(P("margin-top")) ?? 0), Math.Max(0, OdfUnits.Length(P("margin-bottom")) ?? 0)));
        }
    }

    private XElement? EmbeddedListStyle(XElement owner)
    {
        var name = (string?)owner.Attribute(Ns.Presentation + "style-name");
        for (int i = 0; i < 16 && name is not null; i++)
        {
            var style = styles.Find("presentation", name);
            if (style is null) return null;
            if (style.Descendants(Ns.Text + "list-style").FirstOrDefault() is { } list) return list;
            name = (string?)style.Attribute(Ns.Style + "parent-style-name");
        }
        return null;
    }

    private Dictionary<string, string> LevelProps(XElement owner, int level, Dictionary<string, string> g)
    {
        var name = (string?)owner.Attribute(Ns.Presentation + "style-name");
        for (int i = 0; i < 16 && styles.Find("presentation", name) is { } style; i++)
        {
            var parent = (string?)style.Attribute(Ns.Style + "parent-style-name");
            if (parent is null) break;
            name = parent;
        }
        return name is not null && OutlineRoot.Match(name) is { Success: true } m
            ? Merge(g, styles.Properties("presentation", m.Groups[1].Value + Math.Min(level, 9), false))
            : g;
    }

    private static XElement? ListLevel(XElement? listStyle, int level) =>
        listStyle?.Elements().FirstOrDefault(x => (int?)x.Attribute(Ns.Text + "level") == level);

    private static string Label(XElement levelEl, XElement listStyle, int level, int[] counters)
    {
        switch (levelEl.Name.LocalName)
        {
            case "list-level-style-bullet":
            {
                var ch = (string?)levelEl.Attribute(Ns.Text + "bullet-char") ?? "•";
                return ch is "●" or "" ? "•" : ch;
            }
            case "list-level-style-number":
            {
                int display = Math.Clamp((int?)levelEl.Attribute(Ns.Text + "display-levels") ?? 1, 1, level);
                var sb = new StringBuilder();
                for (int l = level - display + 1; l <= level; l++)
                {
                    var format = (string?)(ListLevel(listStyle, l) ?? levelEl).Attribute(Ns.Style + "num-format");
                    if (l > level - display + 1) sb.Append('.');
                    sb.Append(NumberText(Math.Max(1, l < counters.Length ? counters[l] : 1), format));
                }
                return ((string?)levelEl.Attribute(Ns.Style + "num-prefix") ?? "") + sb + ((string?)levelEl.Attribute(Ns.Style + "num-suffix") ?? "");
            }
            default: return "";
        }
    }

    private static string NumberText(int n, string? format) => format switch
    {
        "a" => Letters(n, 'a'),
        "A" => Letters(n, 'A'),
        "i" => Roman(n).ToLowerInvariant(),
        "I" => Roman(n),
        "" or null => "",
        _ => n.ToString(CultureInfo.InvariantCulture),
    };

    private static string Letters(int n, char a)
    {
        var sb = new StringBuilder();
        for (; n > 0 && sb.Length < 8; n = (n - 1) / 26) sb.Insert(0, (char)(a + (n - 1) % 26));
        return sb.ToString();
    }

    private static string Roman(int n)
    {
        if (n is <= 0 or >= 4000) return n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder();
        foreach (var (value, text) in new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") })
            for (; n >= value; n -= value) sb.Append(text);
        return sb.ToString();
    }

    private void Inline(XElement parent, RunProperties run, Hyperlink? link, List<CellTextRun> runs, TextState state)
    {
        foreach (var node in parent.Nodes())
        {
            if (state.Characters > 1_000_000) return;
            if (node is XText text)
            {
                var sb = new StringBuilder();
                foreach (char c in text.Value)
                {
                    if (c is ' ' or '\t' or '\n' or '\r')
                    {
                        if (!state.LastSpace) sb.Append(' ');
                        state.LastSpace = true;
                    }
                    else
                    {
                        sb.Append(c);
                        state.LastSpace = false;
                    }
                }
                Emit(sb.ToString());
            }
            else if (node is XElement e)
            {
                if (e.Name.Namespace == Ns.Presentation)
                {
                    switch (e.Name.LocalName)
                    {
                        case "footer": Emit(_footer); state.LastSpace = false; break;
                        case "date-time": Emit(_date); state.LastSpace = false; break;
                    }
                    continue;
                }
                if (e.Name.Namespace != Ns.Text) continue;
                switch (e.Name.LocalName)
                {
                    case "span":
                        Inline(e, OdfText.Run(styles, styles.Properties("text", (string?)e.Attribute(Ns.Text + "style-name"), false), run), link, runs, state);
                        break;
                    case "a":
                    {
                        var href = (string?)e.Attribute(Ns.XLink + "href");
                        var target = href is not null && (href.Contains("://") || href.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) ? new Hyperlink(href) : link;
                        Inline(e, OdfText.Run(styles, styles.Properties("text", (string?)e.Attribute(Ns.Text + "style-name"), false), run), target, runs, state);
                        break;
                    }
                    case "s":
                        Emit(new string(' ', Math.Clamp((int?)e.Attribute(Ns.Text + "c") ?? 1, 0, 1000)));
                        state.LastSpace = true;
                        break;
                    case "tab": Emit("\t"); state.LastSpace = true; break;
                    case "line-break": Emit("\n"); state.LastSpace = true; break;
                    case "page-number": Emit(_pageNumber.ToString(CultureInfo.InvariantCulture)); state.LastSpace = false; break;
                    case "date" or "time" or "title" or "initial-creator" or "creator" or "file-name" or "page-count" or "sequence" or "bookmark-ref":
                        Inline(e, run, link, runs, state);
                        break;
                    case "meta" or "meta-field" or "ruby" or "ruby-base" or "note-citation" or "change":
                        Inline(e, run, link, runs, state);
                        break;
                }
            }
        }

        void Emit(string value)
        {
            if (value.Length == 0) return;
            state.Characters += value.Length;
            runs.Add(new CellTextRun(value, run, link));
        }
    }

    // ------------------------------------------------------------------ tables

    private sealed record Border(double Width, Color Color);

    private static Border? ParseBorder(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.StartsWith("none", StringComparison.Ordinal) || value.StartsWith("hidden", StringComparison.Ordinal)) return null;
        double width = 0.5;
        Color color = Color.Black;
        foreach (var part in value.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (Color.TryParseHex(part) is { } c) color = c;
            else if (OdfUnits.Length(part) is { } len) width = Math.Max(0.25, len);
        }
        return new Border(width, color);
    }

    private void ReadTable(XElement table, double x, double y, double frameWidth, double frameHeight)
    {
        var columns = new List<(double Width, string? DefaultStyle)>();
        var rows = new List<XElement>();
        var rowHeights = new List<double>();
        void Collect(XElement parent, int depth)
        {
            foreach (var e in parent.Elements())
            {
                if (e.Name.Namespace != Ns.Table || depth > 4) continue;
                switch (e.Name.LocalName)
                {
                    case "table-column":
                    {
                        var props = styles.Properties("table-column", (string?)e.Attribute(Ns.Table + "style-name"), false);
                        double width = OdfUnits.Length(props.GetValueOrDefault("table-column-properties/column-width")) ?? 0;
                        int repeat = Math.Clamp((int?)e.Attribute(Ns.Table + "number-columns-repeated") ?? 1, 1, 2000);
                        for (int i = 0; i < repeat && columns.Count < 2000; i++) columns.Add((width, (string?)e.Attribute(Ns.Table + "default-cell-style-name")));
                        break;
                    }
                    case "table-row":
                    {
                        var props = styles.Properties("table-row", (string?)e.Attribute(Ns.Table + "style-name"), false);
                        double height = OdfUnits.Length(props.GetValueOrDefault("table-row-properties/row-height")) ?? OdfUnits.Length(props.GetValueOrDefault("table-row-properties/min-row-height")) ?? 18;
                        int repeat = Math.Clamp((int?)e.Attribute(Ns.Table + "number-rows-repeated") ?? 1, 1, 5000);
                        for (int i = 0; i < repeat && rows.Count < 5000; i++) { rows.Add(e); rowHeights.Add(height); }
                        break;
                    }
                    case "table-columns" or "table-header-columns" or "table-header-rows" or "table-rows" or "table-row-group" or "table-column-group":
                        Collect(e, depth + 1);
                        break;
                }
            }
        }
        Collect(table, 0);
        if (columns.Count == 0 || rows.Count == 0) return;
        if (columns.Sum(c => c.Width) <= 0)
            for (int i = 0; i < columns.Count; i++) columns[i] = (frameWidth > 0 ? frameWidth / columns.Count : 72, columns[i].DefaultStyle);

        var xs = new double[columns.Count + 1];
        for (int c = 0; c < columns.Count; c++) xs[c + 1] = xs[c] + columns[c].Width;
        var ys = new double[rows.Count + 1];
        for (int r = 0; r < rows.Count; r++) ys[r + 1] = ys[r] + rowHeights[r];
        int id = _tables++;
        TableRowSpan Rows(int first, int count) => new(id, first, count, rowHeights);
        int cells = 0;
        var template = styles.Resource("table-template", (string?)table.Attribute(Ns.Table + "template-name"));
        bool Use(string name) => (string?)table.Attribute(Ns.Table + name) == "true";
        var (firstRow, lastRow, firstColumn, lastColumn) = (Use("use-first-row-styles"), Use("use-last-row-styles"), Use("use-first-column-styles"), Use("use-last-column-styles"));
        bool bandRows = Use("use-banding-rows-styles"), bandColumns = Use("use-banding-columns-styles");
        Dictionary<string, string> TemplateProps(int row, int column)
        {
            var result = new Dictionary<string, string>();
            if (template is null) return result;
            void Layer(string part)
            {
                if (template.Element(Ns.Table + part) is { } e) result = Merge(result, styles.Properties("table-cell", (string?)e.Attribute(Ns.Table + "style-name"), false));
            }
            int body = firstRow ? 1 : 0;
            Layer("body");
            if (bandRows && row >= body && (row - body) % 2 == 0) Layer("odd-rows");
            if (bandColumns && (column - (firstColumn ? 1 : 0)) % 2 == 0) Layer("odd-columns");
            if (lastColumn && column == columns.Count - 1) Layer("last-column");
            if (firstColumn && column == 0) Layer("first-column");
            if (lastRow && row == rows.Count - 1) Layer("last-row");
            if (firstRow && row == 0) Layer("first-row");
            return result;
        }
        var covered = new bool[rows.Count, columns.Count];
        var borders = new List<SheetShape>();

        for (int r = 0; r < rows.Count; r++)
        {
            if (r > 0 && ReferenceEquals(rows[r], rows[r - 1]) && rows[r].Elements(Ns.Table + "table-cell").All(c => !c.HasElements)) { /* repeated empty rows draw nothing but their fills below */ }
            var row = rows[r];
            int col = 0;
            foreach (var cell in row.Elements())
            {
                if (cell.Name.Namespace != Ns.Table || cell.Name.LocalName is not ("table-cell" or "covered-table-cell")) continue;
                int repeat = Math.Clamp((int?)cell.Attribute(Ns.Table + "number-columns-repeated") ?? 1, 1, 2000);
                for (int k = 0; k < repeat && col < columns.Count; k++, col++)
                {
                    if (covered[r, col] || cell.Name.LocalName == "covered-table-cell") continue;
                    if (++cells > MaxTableCells) throw new InvalidDataException("The table has too many cells.");
                    int cs = Math.Clamp((int?)cell.Attribute(Ns.Table + "number-columns-spanned") ?? 1, 1, columns.Count - col);
                    int rs = Math.Clamp((int?)cell.Attribute(Ns.Table + "number-rows-spanned") ?? 1, 1, rows.Count - r);
                    for (int rr = r; rr < r + rs; rr++)
                        for (int cc = col; cc < col + cs; cc++) covered[rr, cc] = true;
                    var styleName = (string?)cell.Attribute(Ns.Table + "style-name") ?? (string?)row.Attribute(Ns.Table + "default-cell-style-name") ?? columns[col].DefaultStyle;
                    var props = Merge(TemplateProps(r, col), styles.Properties("table-cell", styleName, false));
                    double cx = x + xs[col], cy = y + ys[r], cw = xs[col + cs] - xs[col], ch = ys[r + rs] - ys[r];
                    AddCell(cell, props, cx, cy, cw, ch, Rows(r, rs));
                    string? P(string key) => props.GetValueOrDefault("table-cell-properties/" + key);
                    Border? Edge(string side) => ParseBorder(P("border-" + side) ?? P("border"));
                    if (Edge("top") is { } top) borders.Add(Line(cx, cy, cw, 0, top, Rows(r, 0)));
                    if (Edge("bottom") is { } bottom) borders.Add(Line(cx, cy + ch, cw, 0, bottom, Rows(r + rs, 0)));
                    if (Edge("left") is { } left) borders.Add(Line(cx, cy, 0, ch, left, Rows(r, rs)));
                    if (Edge("right") is { } right) borders.Add(Line(cx + cw, cy, 0, ch, right, Rows(r, rs)));
                }
            }
        }
        foreach (var border in borders) _slide.Objects.Add(border with { ZOrder = _z++ });
    }

    private SheetShape Line(double x, double y, double w, double h, Border border, TableRowSpan rows) => new()
    {
        Geometry = ShapeGeometry.Line, Anchor = ImageAnchorKind.Absolute, X = x, Y = y, Width = w, Height = h,
        LineColor = border.Color, LineWidth = border.Width, TableRows = rows,
    };

    private void AddCell(XElement cell, Dictionary<string, string> props, double x, double y, double w, double h, TableRowSpan rows)
    {
        string? P(string key) => props.GetValueOrDefault("table-cell-properties/" + key);
        var g = Merge(styles.Properties("graphic", null), props);
        var paragraphs = ReadText(cell, cell, g, null);
        var fill = props.GetValueOrDefault("graphic-properties/fill") == "solid" ? OdfUnits.Hex(props.GetValueOrDefault("graphic-properties/fill-color")) : OdfUnits.Hex(P("background-color"));
        double Pad(string side, double d) => OdfUnits.Length(P("padding-" + side) ?? P("padding")) ?? d;
        var anchor = P("vertical-align") switch { "middle" => ShapeTextAnchor.Middle, "bottom" => ShapeTextAnchor.Bottom, _ => ShapeTextAnchor.Top };
        if (paragraphs.Count == 0 && fill is null) return;
        _slide.Objects.Add(new SheetShape
        {
            Geometry = ShapeGeometry.Rectangle, Anchor = ImageAnchorKind.Absolute, X = x, Y = y, Width = w, Height = h, ZOrder = _z++,
            Fill = fill, Paragraphs = paragraphs, TextAnchor = anchor, TableRows = rows,
            TextInsets = new Thickness(Pad("top", 3.6), Pad("right", 7.2), Pad("bottom", 3.6), Pad("left", 7.2)),
        });
    }
}
