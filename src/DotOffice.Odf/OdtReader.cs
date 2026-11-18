using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DotOffice.Core;
using DotOffice.Opc;

namespace DotOffice.Odf;

/// <summary>Reads OpenDocument text documents (.odt, .ott) into the <see cref="Document"/> model.</summary>
public static class OdtReader
{
    /// <summary>Reads an ODT package. Malformed, encrypted or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Document Read(Stream stream, OdfReaderOptions? options = null)
    {
        options ??= new OdfReaderOptions();
        try
        {
            using var package = OdfPackage.Open(stream, options, "application/vnd.oasis.opendocument.text", out var styles);
            XElement root;
            using (var r = package.OpenXml("content.xml")) root = XElement.Load(r);
            foreach (var part in root.Elements().Where(e => e.Name.LocalName is "automatic-styles" or "font-face-decls")) styles.Add(part);
            var body = root.Element(Ns.Office + "body")?.Element(Ns.Office + "text")
                       ?? throw new InvalidDataException("The document is not a text document.");
            return new OdtContext(package, styles, options).Read(body);
        }
        catch (System.Xml.XmlException e)
        {
            throw new InvalidDataException($"The document contains malformed XML: {e.Message}", e);
        }
    }
}

internal sealed class OdtContext(OpcPackage package, OdfStyles styles, OdfReaderOptions options)
{
    private const int MaxDepth = 64, MaxBlocks = 2_000_000;
    private static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    private static readonly XNamespace XLink = "http://www.w3.org/1999/xlink";

    private readonly Dictionary<string, int[]> _counters = [];
    private readonly Dictionary<string, RunProperties> _runCache = [];
    private int _blocks, _footnotes;
    private bool _warnedShape;
    private RunProperties _base = RunProperties.Default;

    public Document Read(XElement body)
    {
        var props = styles.Properties("paragraph", null);
        _base = Run(props, new RunProperties { FontFamily = "Liberation Serif", FontSize = 12 });
        var master = styles.MasterPage(FirstMasterName(body)) ?? styles.MasterPage(styles.DefaultMasterPage);
        var layout = styles.PageLayoutOf(master is null ? null : (string?)master.Attribute(Ns.Style + "name"));
        var (header, footer) = (master?.Element(Ns.Style + "header"), master?.Element(Ns.Style + "footer"));

        var lp = layout?.Element(Ns.Style + "page-layout-properties");
        double Len(string a, double d) => OdfUnits.Length((string?)lp?.Attribute(Ns.Fo + a)) ?? d;
        double w = Len("page-width", 612), h = Len("page-height", 792);
        // Border and padding of the page sit inside the margin and take room from the body as well.
        double Inset(string side) => Len("margin-" + side, 56.7) + Len("padding-" + side, 0) + (OdfUnits.Length(((string?)lp?.Attribute(Ns.Fo + "border-" + side) ?? (string?)lp?.Attribute(Ns.Fo + "border"))?.Split(' ')[0]) ?? 0);
        double top = Len("margin-top", 56.7), bottom = Len("margin-bottom", 56.7);
        double mTop = Inset("top"), mBottom = Inset("bottom"), mLeft = Inset("left"), mRight = Inset("right");
        // A page whose margins leave no room for text (a mistake, or a test file) gets them scaled down.
        if (mTop + mBottom > h * 0.8) { double f = h * 0.8 / (mTop + mBottom); mTop *= f; mBottom *= f; top *= f; bottom *= f; }
        if (mLeft + mRight > w * 0.8) { double f = w * 0.8 / (mLeft + mRight); mLeft *= f; mRight *= f; }
        double headerHeight = ExtraHeight(layout, "header-style"), footerHeight = ExtraHeight(layout, "footer-style");
        var section = new Section
        {
            PageWidth = w,
            PageHeight = h,
            Orientation = (string?)lp?.Attribute(Ns.Style + "print-orientation") == "landscape" || w > h ? Orientation.Landscape : Orientation.Portrait,
            Margins = new Thickness(mTop + (header is null ? 0 : headerHeight), mRight, mBottom + (footer is null ? 0 : footerHeight), mLeft),
            HeaderDistance = top,
            FooterDistance = bottom,
            DefaultHeader = header is null ? null : Story(header),
            DefaultFooter = footer is null ? null : Story(footer),
        };
        ReadBlocks(body, section.Blocks, 0);
        var doc = new Document { Settings = new DocumentSettings { DefaultTabStop = OdfUnits.Length(props.GetValueOrDefault("paragraph-properties/tab-stop-distance")) ?? 36 } };
        doc.Sections.Add(section);
        if (section.Blocks.Count == 0) section.Blocks.Add(new Paragraph());
        return doc;
    }

    private string? FirstMasterName(XElement body)
    {
        foreach (var p in body.Descendants().Where(e => e.Name == Ns.Text + "p" || e.Name == Ns.Text + "h").Take(1))
            return styles.Properties("paragraph", (string?)p.Attribute(Ns.Text + "style-name")).GetValueOrDefault("style/master-page-name");
        return null;
    }

    /// <summary>Header or footer content takes room inside the page margin: its minimum height plus the gap below it.</summary>
    private double ExtraHeight(XElement? layout, string kind)
    {
        var p = layout?.Element(Ns.Style + kind)?.Element(Ns.Style + "header-footer-properties");
        if (p is null) return 0;
        return (OdfUnits.Length((string?)p.Attribute(Ns.Fo + "min-height")) ?? 0) + (OdfUnits.Length((string?)p.Attribute(Ns.Fo + "margin-bottom")) ?? 0)
               + (OdfUnits.Length((string?)p.Attribute(Ns.Fo + "margin-top")) ?? 0);
    }

    private HeaderFooter Story(XElement e)
    {
        var story = new HeaderFooter();
        ReadBlocks(e, story.Blocks, 0);
        return story;
    }

    // ------------------------------------------------------------------ blocks

    private void ReadBlocks(XElement parent, List<Block> blocks, int depth, ListContext? list = null)
    {
        if (depth > MaxDepth) throw new InvalidDataException("The document is nested too deeply.");
        foreach (var e in parent.Elements())
        {
            if (++_blocks > MaxBlocks) throw new InvalidDataException("The document has too many blocks.");
            if (e.Name.Namespace != Ns.Text && e.Name.Namespace != Ns.Table) continue;
            switch (e.Name.LocalName)
            {
                case "p" or "h":
                    blocks.Add(ReadParagraph(e, list, depth));
                    break;
                case "list":
                    ReadList(e, blocks, depth, list);
                    break;
                case "table" when e.Name.Namespace == Ns.Table:
                    blocks.Add(ReadTable(e, depth));
                    break;
                case "section" or "index-body" or "table-of-content" or "alphabetical-index" or "illustration-index" or "user-index" or "bibliography":
                    ReadBlocks(e, blocks, depth + 1, list);
                    break;
                case "index-title":
                    ReadBlocks(e, blocks, depth + 1, list);
                    break;
            }
        }
    }

    private sealed class ListContext(string? styleName, int level)
    {
        public string? StyleName { get; } = styleName;
        public int Level { get; } = level;
    }

    private void ReadList(XElement list, List<Block> blocks, int depth, ListContext? outer)
    {
        var styleName = (string?)list.Attribute(Ns.Text + "style-name") ?? outer?.StyleName;
        int level = (outer?.Level ?? 0) + 1;
        if (outer is null && (string?)list.Attribute(Ns.Text + "continue-numbering") != "true" && styleName is not null) _counters.Remove(styleName);
        var context = new ListContext(styleName, level);
        foreach (var item in list.Elements().Where(x => x.Name == Ns.Text + "list-item" || x.Name == Ns.Text + "list-header"))
        {
            bool first = true;
            if (styleName is not null && item.Name.LocalName == "list-item")
            {
                var c = _counters.TryGetValue(styleName, out var arr) ? arr : _counters[styleName] = new int[11];
                if (int.TryParse((string?)item.Attribute(Ns.Text + "start-value"), out var sv)) c[Math.Min(level, 10)] = sv - 1;
                c[Math.Min(level, 10)]++;
                for (int i = level + 1; i <= 10; i++) c[i] = 0;
            }
            foreach (var child in item.Elements())
            {
                if (child.Name == Ns.Text + "p" || child.Name == Ns.Text + "h")
                {
                    blocks.Add(ReadParagraph(child, first && item.Name.LocalName == "list-item" ? context : context, depth, first && item.Name.LocalName == "list-item"));
                    first = false;
                }
                else ReadBlocks(new XElement("x", child), blocks, depth + 1, context);
            }
        }
    }

    private Paragraph ReadParagraph(XElement e, ListContext? list, int depth, bool labelled = false)
    {
        var styleName = (string?)e.Attribute(Ns.Text + "style-name");
        var p = styles.Properties("paragraph", styleName);
        var run = Run(p, _base);
        var inlines = new List<Inline>();
        var state = new TextState { Space = true };
        ReadInlines(e, run, inlines, state, null, depth + 1);

        var props = ParagraphProps(p, styleName);
        ListLabel? label = null;
        if (list?.StyleName is not null && labelled && styles.ListStyle(list.StyleName) is { } ls)
            (props, label) = ApplyList(props, ls, list, run);
        var paragraph = new Paragraph { Properties = props, MarkProperties = run, ListLabel = label };
        paragraph.Inlines.AddRange(inlines);
        return paragraph;
    }

    private (ParagraphProperties, ListLabel?) ApplyList(ParagraphProperties props, XElement listStyle, ListContext list, RunProperties run)
    {
        int level = Math.Min(list.Level, 10);
        var levelEl = listStyle.Elements().FirstOrDefault(x => (int?)x.Attribute(Ns.Text + "level") == level);
        if (levelEl is null) return (props, null);
        var lp = levelEl.Element(Ns.Style + "list-level-properties");
        var la = lp?.Element(Ns.Style + "list-level-label-alignment");
        double left, indent;
        if (la is not null)
        {
            left = OdfUnits.Length((string?)la.Attribute(Ns.Fo + "margin-left")) ?? 0;
            indent = OdfUnits.Length((string?)la.Attribute(Ns.Fo + "text-indent")) ?? 0;
        }
        else
        {
            double before = OdfUnits.Length((string?)lp?.Attribute(Ns.Text + "space-before")) ?? 0;
            double width = OdfUnits.Length((string?)lp?.Attribute(Ns.Text + "min-label-width")) ?? 0;
            left = before + width;
            indent = -width;
        }
        var tp = levelEl.Element(Ns.Style + "text-properties");
        var labelRun = run;
        if (tp is not null) labelRun = Run(Flatten(tp), run);
        string text;
        if (levelEl.Name.LocalName == "list-level-style-bullet")
            text = (string?)levelEl.Attribute(Ns.Text + "bullet-char") ?? "•";
        else if (levelEl.Name.LocalName == "list-level-style-number")
        {
            var counters = _counters[list.StyleName!];
            int display = Math.Clamp((int?)levelEl.Attribute(Ns.Text + "display-levels") ?? 1, 1, level);
            var sb = new StringBuilder();
            for (int l = level - display + 1; l <= level; l++)
            {
                var format = (string?)(listStyle.Elements().FirstOrDefault(x => (int?)x.Attribute(Ns.Text + "level") == l) ?? levelEl).Attribute(Styles("num-format"));
                if (l > level - display + 1) sb.Append('.');
                sb.Append(Format(counters[l] == 0 && l != level ? 1 : counters[l], format));
            }
            text = ((string?)levelEl.Attribute(Styles("num-prefix")) ?? "") + sb + ((string?)levelEl.Attribute(Styles("num-suffix")) ?? "");
        }
        else return (props, null);
        props = props with { LeftIndent = left, FirstLineIndent = indent };
        if (text.Length == 0) return (props, null);
        var suffix = (string?)la?.Attribute(Ns.Text + "label-followed-by") switch { "space" => ListLabelSuffix.Space, "nothing" => ListLabelSuffix.Nothing, _ => ListLabelSuffix.Tab };
        return (props, new ListLabel(text, labelRun with { Underline = UnderlineStyle.None }, suffix));
    }

    private static XName Styles(string name) => Ns.Style + name;

    private static string Format(int n, string? format) => format switch
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
        for (; n > 0; n = (n - 1) / 26) sb.Insert(0, (char)(a + (n - 1) % 26));
        return sb.ToString();
    }

    private static string Roman(int n)
    {
        if (n is <= 0 or > 3999) return n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder();
        int[] v = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
        string[] s = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
        for (int i = 0; i < v.Length; i++)
            while (n >= v[i]) { sb.Append(s[i]); n -= v[i]; }
        return sb.ToString();
    }

    private static Dictionary<string, string> Flatten(XElement props)
    {
        var d = new Dictionary<string, string>();
        foreach (var a in props.Attributes())
            if (!a.IsNamespaceDeclaration) d[props.Name.LocalName + "/" + a.Name.LocalName] = a.Value;
        return d;
    }

    // ------------------------------------------------------------------ inlines

    private sealed class TextState
    {
        /// <summary>The previous character was a space (or the paragraph start), so a further white space character is dropped.</summary>
        public bool Space { get; set; }
    }

    private void ReadInlines(XElement parent, RunProperties run, List<Inline> inlines, TextState state, Hyperlink? link, int depth)
    {
        if (depth > MaxDepth) throw new InvalidDataException("The document is nested too deeply.");
        foreach (var node in parent.Nodes())
        {
            if (node is XText text)
            {
                var t = Collapse(text.Value, state);
                if (t.Length > 0) Add(inlines, new TextRun(t, run), link);
                continue;
            }
            if (node is not XElement e || e.Name.Namespace != Ns.Text && e.Name.Namespace != Draw) continue;
            switch (e.Name.LocalName)
            {
                case "span":
                {
                    var r = Run(styles.Properties("text", (string?)e.Attribute(Ns.Text + "style-name")), run);
                    ReadInlines(e, r, inlines, state, link, depth + 1);
                    break;
                }
                case "a":
                {
                    var href = (string?)e.Attribute(XLink + "href");
                    var l = href is null ? link : href.StartsWith('#') ? new Hyperlink(null, href[1..]) : new Hyperlink(href);
                    var r = Run(styles.Properties("text", (string?)e.Attribute(Ns.Text + "style-name")), run);
                    ReadInlines(e, r, inlines, state, l, depth + 1);
                    break;
                }
                case "s":
                    Add(inlines, new TextRun(new string(' ', Math.Clamp((int?)e.Attribute(Ns.Text + "c") ?? 1, 1, 1000)), run), link);
                    state.Space = true;
                    break;
                case "tab":
                    Add(inlines, new Tab(run), link);
                    state.Space = true;
                    break;
                case "line-break":
                    Add(inlines, new Break(BreakKind.Line, run), link);
                    state.Space = true;
                    break;
                case "bookmark" or "bookmark-start":
                    if ((string?)e.Attribute(Ns.Text + "name") is { } name) inlines.Add(new BookmarkStart(name, run));
                    break;
                case "page-number":
                    Add(inlines, new FieldRun(FieldKind.Page, run) { CachedResult = e.Value }, link);
                    state.Space = false;
                    break;
                case "page-count":
                    Add(inlines, new FieldRun(FieldKind.NumPages, run) { CachedResult = e.Value }, link);
                    state.Space = false;
                    break;
                case "note":
                    ReadNote(e, run, inlines, depth);
                    break;
                case "frame":
                    ReadFrame(e, run, inlines, depth);
                    break;
                case "tracked-changes" or "note-body" or "note-citation" or "annotation":
                    break;
                default:
                    // Fields, ruby, change marks and the like: their text is the cached result.
                    ReadInlines(e, run, inlines, state, link, depth + 1);
                    break;
            }
        }
    }

    private static void Add(List<Inline> inlines, Inline inline, Hyperlink? link)
    {
        inline.Link = link;
        inlines.Add(inline);
    }

    /// <summary>ODF white space handling: runs of space, tab and line feed characters collapse to one space, and a space after another one (or at the start) is dropped.</summary>
    private static string Collapse(string value, TextState state)
    {
        var sb = new StringBuilder(value.Length);
        foreach (char c in value)
        {
            if (c is ' ' or '\t' or '\n' or '\r')
            {
                if (state.Space) continue;
                sb.Append(' ');
                state.Space = true;
            }
            else
            {
                sb.Append(c);
                state.Space = false;
            }
        }
        return sb.ToString();
    }

    private void ReadNote(XElement e, RunProperties run, List<Inline> inlines, int depth)
    {
        var kind = (string?)e.Attribute(Ns.Text + "note-class") == "endnote" ? NoteKind.Endnote : NoteKind.Footnote;
        var citation = e.Element(Ns.Text + "note-citation");
        var note = new Note(kind) { Mark = citation?.Value is { Length: > 0 } c && !int.TryParse(c, out _) ? "" : (++_footnotes).ToString(CultureInfo.InvariantCulture) };
        var sup = run with { VerticalPosition = VerticalPosition.Superscript };
        if (e.Element(Ns.Text + "note-body") is { } body) ReadBlocks(body, note.Blocks, depth + 1);
        if (note.Blocks.FirstOrDefault() is Paragraph first) first.Inlines.Insert(0, new NoteMark(note, sup));
        inlines.Add(new NoteReference(note, sup));
    }

    private void ReadFrame(XElement frame, RunProperties run, List<Inline> inlines, int depth)
    {
        double w = OdfUnits.Length((string?)frame.Attribute(Ns.Svg + "width")) ?? 0, h = OdfUnits.Length((string?)frame.Attribute(Ns.Svg + "height")) ?? 0;
        var g = styles.Properties("graphic", (string?)frame.Attribute(Draw + "style-name"));
        bool inline = (string?)frame.Attribute(Ns.Text + "anchor-type") == "as-char";
        double x = OdfUnits.Length((string?)frame.Attribute(Ns.Svg + "x")) ?? 0, y = OdfUnits.Length((string?)frame.Attribute(Ns.Svg + "y")) ?? 0;
        string anchor = (string?)frame.Attribute(Ns.Text + "anchor-type") ?? "paragraph";
        var horizontal = new FloatPosition(anchor == "page" ? FloatReference.Page : FloatReference.Column, FloatAlignment.Offset, x);
        var vertical = new FloatPosition(anchor == "page" ? FloatReference.Page : FloatReference.Paragraph, FloatAlignment.Offset, y);
        var wrap = g.GetValueOrDefault("graphic-properties/wrap") switch { "none" => TextWrap.TopAndBottom, "run-through" => TextWrap.None, _ => TextWrap.Square };
        bool behind = g.GetValueOrDefault("graphic-properties/run-in-background") == "true";

        if (frame.Element(Draw + "image") is { } image)
        {
            var data = LoadImage((string?)image.Attribute(XLink + "href"), image);
            if (data is null) return;
            if (w <= 0 || h <= 0) return;
            if (inline) inlines.Add(new ImageRun(data, w, h, run));
            else inlines.Add(new FloatingImage(data, w, h, run) { Horizontal = horizontal, Vertical = vertical, Wrap = wrap, BehindText = behind });
        }
        else if (frame.Element(Draw + "text-box") is { } box)
        {
            var tb = new FloatingTextBox(w, h > 0 ? h : 20, run)
            {
                Horizontal = horizontal, Vertical = vertical, Wrap = inline ? TextWrap.None : wrap, BehindText = behind,
                Fill = OdfUnits.Hex(g.GetValueOrDefault("graphic-properties/fill-color")) is { } fc && g.GetValueOrDefault("graphic-properties/fill") != "none" ? fc : null,
                Insets = new Thickness(OdfUnits.Length(g.GetValueOrDefault("graphic-properties/padding-top")) ?? 0, OdfUnits.Length(g.GetValueOrDefault("graphic-properties/padding-right")) ?? 0,
                    OdfUnits.Length(g.GetValueOrDefault("graphic-properties/padding-bottom")) ?? 0, OdfUnits.Length(g.GetValueOrDefault("graphic-properties/padding-left")) ?? 0),
            };
            ReadBlocks(box, tb.Blocks, depth + 1);
            inlines.Add(inline ? new InlineShape(tb) : tb);
        }
        else if (!_warnedShape)
        {
            _warnedShape = true;
            options.Warnings.Warn("odt.object-skipped", "A drawing, chart or embedded object in the document is not drawn.");
        }
    }

    private ImageData? LoadImage(string? href, XElement image)
    {
        if (href is not null && !href.Contains("://") && !href.StartsWith('/') && !href.Contains("..") && package.Exists(href))
            return new ImageData(package.ReadAllBytes(href), href);
        if (image.Element(Ns.Office + "binary-data") is { } b)
        {
            try { return new ImageData(Convert.FromBase64String(b.Value)); }
            catch (FormatException) { return null; }
        }
        options.Warnings.Warn("odt.image-missing", $"The image '{href}' is not in the package or is external; it is not drawn.");
        return null;
    }

    // ------------------------------------------------------------------ tables

    private Table ReadTable(XElement t, int depth)
    {
        if (depth > MaxDepth) throw new InvalidDataException("The document is nested too deeply.");
        var tp = styles.Properties("table", (string?)t.Attribute(Ns.Table + "style-name"));
        double? width = OdfUnits.Length(tp.GetValueOrDefault("table-properties/width"));
        var table = new Table
        {
            Properties = new TableProperties
            {
                Width = width is { } wd ? new TableWidth(wd, TableWidthUnit.Points) : TableWidth.Auto,
                Alignment = tp.GetValueOrDefault("table-properties/align") switch { "center" => TableAlignment.Center, "right" => TableAlignment.Right, _ => TableAlignment.Left },
                Indent = OdfUnits.Length(tp.GetValueOrDefault("table-properties/margin-left")) ?? 0,
                Layout = TableLayoutMode.Fixed,
            },
        };
        foreach (var col in Descendants(t, "table-column"))
        {
            int n = Math.Clamp((int?)col.Attribute(Ns.Table + "number-columns-repeated") ?? 1, 1, 1000);
            double cw = OdfUnits.Length(styles.Properties("table-column", (string?)col.Attribute(Ns.Table + "style-name")).GetValueOrDefault("table-column-properties/column-width")) ?? 72;
            for (int i = 0; i < n && table.Grid.Count < 1000; i++) table.Grid.Add(cw);
        }
        var pending = new int[Math.Max(table.Grid.Count, 1) + 1000];
        var pendingSpan = new int[pending.Length];
        foreach (var (rowEl, header) in Rows(t))
        {
            var rp = styles.Properties("table-row", (string?)rowEl.Attribute(Ns.Table + "style-name"));
            double? rh = OdfUnits.Length(rp.GetValueOrDefault("table-row-properties/row-height")) ?? OdfUnits.Length(rp.GetValueOrDefault("table-row-properties/min-row-height"));
            var row = new TableRow
            {
                Height = rh ?? 0,
                HeightRule = rh is null ? RowHeightRule.Auto : rp.ContainsKey("table-row-properties/row-height") ? RowHeightRule.Exact : RowHeightRule.AtLeast,
                IsHeader = header,
            };
            int col = 0;
            foreach (var c in rowEl.Elements())
            {
                if (col >= pending.Length - 1) break;
                int repeat = Math.Clamp((int?)c.Attribute(Ns.Table + "number-columns-repeated") ?? 1, 1, 1000);
                if (c.Name == Ns.Table + "covered-table-cell")
                {
                    for (int i = 0; i < repeat && col < pending.Length - 1; i++)
                    {
                        if (pending[col] > 0)
                        {
                            int span = Math.Max(pendingSpan[col], 1);
                            row.Cells.Add(new TableCell { GridSpan = span, VerticalMerge = VerticalMerge.Continue, Blocks = { new Paragraph() } });
                            for (int k = 0; k < span; k++) pending[col + k]--;
                            col += span;
                        }
                        else col++;
                    }
                    continue;
                }
                if (c.Name != Ns.Table + "table-cell") continue;
                int colSpan = Math.Clamp((int?)c.Attribute(Ns.Table + "number-columns-spanned") ?? 1, 1, 1000);
                int rowSpan = Math.Clamp((int?)c.Attribute(Ns.Table + "number-rows-spanned") ?? 1, 1, 100000);
                for (int i = 0; i < repeat && col < pending.Length - 1 - colSpan; i++)
                {
                    row.Cells.Add(ReadCell(c, colSpan, rowSpan > 1, table, col, depth));
                    if (rowSpan > 1)
                        for (int k = 0; k < colSpan; k++) { pending[col + k] = rowSpan - 1; pendingSpan[col + k] = colSpan; }
                    col += colSpan;
                }
            }
            table.Rows.Add(row);
            if (table.Rows.Count > 100_000) throw new InvalidDataException("The table has too many rows.");
        }
        if (table.Grid.Count == 0)
            for (int i = 0; i < table.Rows.Max(r => r.Cells.Sum(c => c.GridSpan)) ; i++) table.Grid.Add(72);
        return table;
    }

    private static IEnumerable<XElement> Descendants(XElement table, string name)
    {
        foreach (var e in table.Elements())
        {
            if (e.Name == Ns.Table + name) yield return e;
            else if (e.Name.Namespace == Ns.Table && e.Name.LocalName is "table-columns" or "table-header-columns" or "table-column-group")
                foreach (var d in Descendants(e, name)) yield return d;
        }
    }

    private static IEnumerable<(XElement Row, bool Header)> Rows(XElement table, bool header = false)
    {
        foreach (var e in table.Elements())
        {
            if (e.Name == Ns.Table + "table-row")
            {
                int n = Math.Clamp((int?)e.Attribute(Ns.Table + "number-rows-repeated") ?? 1, 1, 1000);
                for (int i = 0; i < n; i++) yield return (e, header);
            }
            else if (e.Name.Namespace == Ns.Table && e.Name.LocalName is "table-rows" or "table-row-group")
                foreach (var r in Rows(e, header)) yield return r;
            else if (e.Name == Ns.Table + "table-header-rows")
                foreach (var r in Rows(e, true)) yield return r;
        }
    }

    private TableCell ReadCell(XElement c, int colSpan, bool restart, Table table, int col, int depth)
    {
        var p = styles.Properties("table-cell", (string?)c.Attribute(Ns.Table + "style-name"));
        double Pad(string side, double fallback) =>
            OdfUnits.Length(p.GetValueOrDefault("table-cell-properties/padding-" + side) ?? p.GetValueOrDefault("table-cell-properties/padding")) ?? fallback;
        double width = 0;
        for (int i = col; i < col + colSpan && i < table.Grid.Count; i++) width += table.Grid[i];
        var cell = new TableCell
        {
            Width = width > 0 ? new TableWidth(width, TableWidthUnit.Points) : TableWidth.Auto,
            GridSpan = colSpan,
            VerticalMerge = restart ? VerticalMerge.Restart : VerticalMerge.None,
            Shading = OdfUnits.Hex(p.GetValueOrDefault("table-cell-properties/background-color")),
            Borders = new BorderSet { Top = Edge(p, "top"), Bottom = Edge(p, "bottom"), Left = Edge(p, "left"), Right = Edge(p, "right") },
            Margins = new Thickness(Pad("top", 0), Pad("right", 5.4), Pad("bottom", 0), Pad("left", 5.4)),
            VerticalAlignment = p.GetValueOrDefault("table-cell-properties/vertical-align") switch
            {
                "middle" => CellVerticalAlignment.Center,
                "bottom" => CellVerticalAlignment.Bottom,
                _ => CellVerticalAlignment.Top,
            },
        };
        ReadBlocks(c, cell.Blocks, depth + 1);
        if (cell.Blocks.Count == 0) cell.Blocks.Add(new Paragraph());
        return cell;
    }

    // ------------------------------------------------------------------ properties

    private static Border Edge(Dictionary<string, string> p, string side, string element = "table-cell-properties")
    {
        var value = p.GetValueOrDefault(element + "/border-" + side) ?? p.GetValueOrDefault(element + "/border");
        if (value is null or "none") return Border.None;
        double width = 0.75;
        var kind = BorderStyle.Single;
        Color? color = null;
        foreach (var part in value.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (char.IsDigit(part[0]) && OdfUnits.Length(part) is { } w) width = w;
            else if (part == "double") kind = BorderStyle.Double;
            else if (part == "dotted") kind = BorderStyle.Dotted;
            else if (part == "dashed") kind = BorderStyle.Dashed;
            else if (part == "hidden") return Border.None;
            else if (Color.TryParseHex(part) is { } c) color = c;
        }
        return new Border(kind, width, color ?? Color.Black);
    }

    private ParagraphProperties ParagraphProps(Dictionary<string, string> p, string? styleName)
    {
        double L(string key) => OdfUnits.Length(p.GetValueOrDefault("paragraph-properties/" + key)) ?? 0;
        var spacing = LineSpacing.Single;
        if (p.GetValueOrDefault("paragraph-properties/line-height") is { } lh)
        {
            if (lh.EndsWith('%') && double.TryParse(lh.AsSpan(0, lh.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct) && pct > 0)
                spacing = new LineSpacing(pct / 100, LineSpacingRule.Auto);
            else if (OdfUnits.Length(lh) is { } len and > 0) spacing = new LineSpacing(len, LineSpacingRule.Exact);
        }
        else if (OdfUnits.Length(p.GetValueOrDefault("paragraph-properties/line-height-at-least")) is { } least and > 0)
            spacing = new LineSpacing(least, LineSpacingRule.AtLeast);
        var rtl = p.GetValueOrDefault("paragraph-properties/writing-mode") is "rl-tb" or "rl";
        var align = p.GetValueOrDefault("paragraph-properties/text-align") switch
        {
            "center" => ParagraphAlignment.Center,
            "end" or "right" => rtl ? ParagraphAlignment.Left : ParagraphAlignment.Right,
            "justify" => ParagraphAlignment.Justify,
            _ => ParagraphAlignment.Left,
        };
        return new ParagraphProperties
        {
            StyleId = styleName,
            Alignment = align,
            LeftIndent = L("margin-left"),
            RightIndent = L("margin-right"),
            FirstLineIndent = L("text-indent"),
            SpaceBefore = L("margin-top"),
            SpaceAfter = L("margin-bottom"),
            LineSpacing = spacing,
            KeepWithNext = p.GetValueOrDefault("paragraph-properties/keep-with-next") == "always",
            KeepLinesTogether = p.GetValueOrDefault("paragraph-properties/keep-together") == "always",
            PageBreakBefore = p.GetValueOrDefault("paragraph-properties/break-before") == "page",
            Shading = OdfUnits.Hex(p.GetValueOrDefault("paragraph-properties/background-color")),
            RightToLeft = rtl,
            Borders = new BorderSet
            {
                Top = Edge(p, "top", "paragraph-properties"), Bottom = Edge(p, "bottom", "paragraph-properties"),
                Left = Edge(p, "left", "paragraph-properties"), Right = Edge(p, "right", "paragraph-properties"),
            },
        };
    }

    /// <summary>Character formatting from the <c>text-properties</c> of a style, on top of <paramref name="inherited"/>.</summary>
    private RunProperties Run(Dictionary<string, string> p, RunProperties inherited)
    {
        string? V(string key) => p.GetValueOrDefault("text-properties/" + key);
        var r = inherited;
        if ((V("font-family") ?? V("font-name")) is { } family) r = r with { FontFamily = styles.FontFamily(family) };
        if ((V("font-name-asian") ?? V("font-family-asian")) is { } asian) r = r with { EastAsianFontFamily = styles.FontFamily(asian) };
        if ((V("font-name-complex") ?? V("font-family-complex")) is { } complex) r = r with { ComplexScriptFontFamily = styles.FontFamily(complex) };
        if (V("font-size") is { } fs)
        {
            if (fs.EndsWith('%') && double.TryParse(fs.AsSpan(0, fs.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct) && pct > 0)
                r = r with { FontSize = inherited.FontSize * pct / 100 };
            else if (OdfUnits.Length(fs) is { } size and > 0) r = r with { FontSize = size };
        }
        if (OdfUnits.Length(V("font-size-complex")) is { } cs and > 0) r = r with { ComplexScriptFontSize = cs };
        if (V("font-weight") is { } w) r = r with { Bold = w == "bold" || int.TryParse(w, out var n) && n >= 600 };
        if (V("font-style") is { } st) r = r with { Italic = st is "italic" or "oblique" };
        if (V("text-underline-style") is { } us) r = r with { Underline = us == "none" ? UnderlineStyle.None : V("text-underline-type") == "double" ? UnderlineStyle.Double : UnderlineStyle.Single };
        if (V("text-line-through-style") is { } ls) r = r with { Strike = ls != "none" };
        if (V("color") is { } color) r = r with { Color = OdfUnits.Hex(color) };
        if (V("background-color") is { } bg) r = r with { Highlight = OdfUnits.Hex(bg) };
        if (V("font-variant") is { } fv) r = r with { SmallCaps = fv == "small-caps" };
        if (V("text-transform") is { } tt) r = r with { Caps = tt == "uppercase" };
        if (V("text-display") is { } td) r = r with { Hidden = td == "none" };
        if (V("language") is { } lang) r = r with { Language = V("country") is { } country && country != "none" ? lang + "-" + country : lang };
        if (OdfUnits.Length(V("letter-spacing")) is { } ls2) r = r with { CharacterSpacing = ls2 };
        if (V("text-position") is { } pos)
        {
            var first = pos.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
            r = r with
            {
                VerticalPosition = first.StartsWith("sub", StringComparison.Ordinal) || first.StartsWith('-') ? VerticalPosition.Subscript
                    : first == "0%" ? VerticalPosition.Baseline : VerticalPosition.Superscript,
            };
            if (pos.Split(' ', StringSplitOptions.RemoveEmptyEntries) is { Length: > 1 } parts && parts[1].EndsWith('%')
                && double.TryParse(parts[1].AsSpan(0, parts[1].Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var scale) && scale > 0 && scale < 100)
                r = r with { FontSize = r.FontSize * scale / 100 };
        }
        return r;
    }
}
