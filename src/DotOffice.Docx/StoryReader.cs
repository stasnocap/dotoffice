using System.Globalization;
using System.Text;
using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Table-style formatting that applies to paragraphs inside one cell.</summary>
internal sealed record CellStyle(ParaFormat? Para, RunFormat? Run);

/// <summary>
/// Reads one story (main document body, header or footer): blocks, paragraphs, runs, fields,
/// inline images and tables. Unknown elements are skipped; nothing here throws on odd content.
/// </summary>
internal sealed class StoryReader(DocxContext context, string partName)
{
    private readonly Stack<FieldFrame> _fields = new();
    private readonly Dictionary<(string?, RunFormat?, RunFormat?), RunProperties> _runCache = [];

    private StyleSheet Styles => context.Styles;

    private readonly Stack<Hyperlink> _links = new();
    private readonly Dictionary<string, string> _openBookmarks = new(StringComparer.Ordinal);

    private sealed class FieldFrame
    {
        public StringBuilder Instruction { get; } = new();
        public bool Separated { get; set; }
        public FieldInstruction? Field { get; set; }
        public RunProperties? Properties { get; init; }

        /// <summary>The run that replaces the cached result of a field we evaluate ourselves.</summary>
        public Inline? Emitted { get; set; }

        /// <summary>Some cached result content was shown.</summary>
        public bool HasResult { get; set; }

        public void Separate()
        {
            if (Separated) return;
            Separated = true;
            Field = FieldInstruction.Parse(Instruction.ToString());
        }
    }

    private sealed class BlockSink(List<Block> target, Func<SectionData, List<Block>>? onSection)
    {
        public List<Block> Target { get; private set; } = target;

        public void EndSection(SectionData data)
        {
            if (onSection is not null) Target = onSection(data);
        }
    }

    /// <summary>Paragraph-level state while reading its runs.</summary>
    private sealed class ParagraphState(Paragraph paragraph, string? styleId, CellStyle? cell)
    {
        public Paragraph Paragraph { get; } = paragraph;
        public string? StyleId { get; } = styleId;
        public CellStyle? Cell { get; } = cell;
    }

    public void ReadBlocks(XmlReader r, List<Block> target, CellStyle? cell, int nesting, Func<SectionData, List<Block>>? onSection) =>
        ReadBlocks(r, new BlockSink(target, onSection), cell, nesting);

    private void ReadBlocks(XmlReader r, BlockSink sink, CellStyle? cell, int nesting)
    {
        int depth = r.Depth;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (r.Depth > context.Options.MaxXmlDepth)
            {
                context.WarnOnce("docx.depth", "Content nested too deeply was skipped.");
                r.Skip();
                continue;
            }
            if (r.IsMC() && r.LocalName == "AlternateContent")
            {
                ReadAlternateContent(r, () => ReadBlocks(r, sink, cell, nesting));
                continue;
            }
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "p":
                    SectionData? sectionEnd = null;
                    var paragraph = ReadParagraph(r, cell, data => sectionEnd = data);
                    sink.Target.Add(paragraph);
                    if (sectionEnd is not null) sink.EndSection(sectionEnd);
                    break;
                case "tbl":
                    if (ReadTable(r, nesting + 1) is { } table) sink.Target.Add(table);
                    break;
                case "sdt":
                    ReadContainer(r, "sdtContent", () => ReadBlocks(r, sink, cell, nesting));
                    break;
                case "customXml":
                case "ins":
                case "moveTo":
                    ReadBlocks(r, sink, cell, nesting);
                    break;
                case "sectPr":
                    sink.EndSection(ReadSection(r));
                    break;
                case "altChunk":
                    context.WarnOnce("docx.altchunk", "Embedded alternative-format chunks (altChunk) are not supported.");
                    r.Skip();
                    break;
                default:
                    r.Skip();
                    break;
            }
        }
    }

    /// <summary>Calls <paramref name="readContent"/> positioned on the named child (e.g. w:sdtContent), skipping the rest.</summary>
    private static void ReadContainer(XmlReader r, string contentName, Action readContent)
    {
        int depth = r.Depth;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (r.IsW() && r.LocalName == contentName) readContent();
            else r.Skip();
        }
    }

    /// <summary>mc:AlternateContent: reads the first mc:Choice (or the fallback when there is none).</summary>
    private static void ReadAlternateContent(XmlReader r, Action readContent)
    {
        int depth = r.Depth;
        bool done = false;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (!done && r.IsMC() && r.LocalName is "Choice" or "Fallback")
            {
                done = true;
                readContent();
            }
            else r.Skip();
        }
    }

    // ---------------------------------------------------------------- paragraphs

    private Paragraph ReadParagraph(XmlReader r, CellStyle? cell, Action<SectionData> onSection)
    {
        int depth = r.Depth;
        // After this, 'open' means the reader sits inside the paragraph (on its first child element).
        bool open = r.Enter() && r.NextChild(depth);
        var direct = ParaFormat.Empty;
        RunFormat? mark = null;
        if (open && r.IsW() && r.LocalName == "pPr")
            direct = ParaFormat.Parse(r, m => mark = m, s => onSection(ReadSection(s)));

        var styleId = direct.StyleId is { } id && Styles.Get(id)?.Type == StyleType.Paragraph ? id : Styles.DefaultParagraphStyle;
        var styleFormat = Styles.Resolve(styleId).Para;
        int? numId = direct.NumId ?? styleFormat.NumId;
        bool numberingIsDirect = direct.NumId is not null;
        NumberingLevel? level = null;
        int ilvl = 0;
        if (numId is > 0)
        {
            ilvl = direct.Ilvl ?? context.Numbering.FindLevelForStyle(numId.Value, styleId, Styles) ?? styleFormat.Ilvl ?? 0;
            level = context.Numbering.GetLevel(numId.Value, ilvl, Styles);
        }

        var format = Styles.ResolveParagraphFormat(cell?.Para, styleId, level?.Para, numberingIsDirect, direct);
        var markProperties = ResolveRun(styleId, cell, mark);
        ListLabel? label = null;
        if (level is not null && context.Numbering.Next(numId!.Value, ilvl, Styles) is { } numbering)
            label = MakeLabel(numbering, styleId, cell, mark);

        var paragraph = new Paragraph
        {
            Properties = StyleSheet.ToProperties(format),
            MarkProperties = markProperties,
            ListLabel = label,
        };

        if (open)
        {
            var state = new ParagraphState(paragraph, styleId, cell);
            // The first child was already consumed if it was pPr; otherwise the reader sits on it.
            if (r.Depth > depth && r.NodeType == XmlNodeType.Element) ReadInlineElement(r, state);
            ReadInlines(r, depth, state);
        }
        MergeAdjacentText(paragraph.Inlines);
        return paragraph;
    }

    private ListLabel MakeLabel(NumberingResult numbering, string? styleId, CellStyle? cell, RunFormat? mark)
    {
        var markFormat = mark ?? RunFormat.Empty;
        var properties = ResolveRun(styleId, cell, markFormat.Merge(numbering.Level.Run));
        var text = numbering.Text;
        if (SymbolFonts.IsSymbolFont(properties.FontFamily))
        {
            text = SymbolFonts.Map(properties.FontFamily, text, unknown: '•');
            // Draw the mapped Unicode symbol with the paragraph's text font instead of the symbol font.
            properties = properties with { FontFamily = ResolveRun(styleId, cell, markFormat).FontFamily };
        }
        return new ListLabel(text, properties, numbering.Level.Suffix, numbering.Level.Alignment);
    }

    private RunProperties ResolveRun(string? styleId, CellStyle? cell, RunFormat? direct)
    {
        var key = (styleId, cell?.Run, direct);
        if (!_runCache.TryGetValue(key, out var props))
            _runCache[key] = props = Styles.ResolveRun(cell?.Run, styleId, direct);
        return props;
    }

    private void ReadInlines(XmlReader r, int depth, ParagraphState state)
    {
        while (r.NextChild(depth)) ReadInlineElement(r, state);
    }

    private void ReadInlineElement(XmlReader r, ParagraphState state)
    {
        if (r.Depth > context.Options.MaxXmlDepth)
        {
            context.WarnOnce("docx.depth", "Content nested too deeply was skipped.");
            r.Skip();
            return;
        }
        if (r.IsMC() && r.LocalName == "AlternateContent")
        {
            ReadAlternateContent(r, () =>
            {
                int d = r.Depth;
                if (r.Enter()) ReadInlines(r, d, state);
            });
            return;
        }
        if (!r.IsW())
        {
            if (r.LocalName is "oMath" or "oMathPara")
                context.WarnOnce("docx.math", "Equations (OMML) are not supported and were skipped.");
            r.Skip();
            return;
        }
        int depth = r.Depth;
        switch (r.LocalName)
        {
            case "r":
                ReadRun(r, state);
                break;
            case "hyperlink":
                var link = HyperlinkTarget(r);
                if (link is not null) _links.Push(link);
                if (r.Enter()) ReadInlines(r, depth, state);
                if (link is not null) _links.Pop();
                break;
            case "bookmarkStart":
                if (r.WAttr("name") is { Length: > 0 } name)
                {
                    if (r.WAttr("id") is { } bookmarkId) _openBookmarks[bookmarkId] = name;
                    context.Bookmarks.TryAdd(name, new StringBuilder());
                    AddInline(state, new BookmarkStart(name, state.Paragraph.MarkProperties));
                }
                r.Skip();
                break;
            case "bookmarkEnd":
                if (r.WAttr("id") is { } endId) _openBookmarks.Remove(endId);
                r.Skip();
                break;
            case "smartTag":
            case "customXml":
            case "ins":
            case "moveTo":
            case "dir":
            case "bdo":
                if (r.Enter()) ReadInlines(r, depth, state);
                break;
            case "sdt":
                ReadContainer(r, "sdtContent", () =>
                {
                    int d = r.Depth;
                    if (r.Enter()) ReadInlines(r, d, state);
                });
                break;
            case "fldSimple":
                var frame = new FieldFrame { Properties = state.Paragraph.MarkProperties };
                frame.Instruction.Append(r.WAttr("instr"));
                frame.Separate();
                _fields.Push(frame);
                if (r.Enter()) ReadInlines(r, depth, state);
                EndField(state);
                break;
            default:
                r.Skip(); // pPr duplicates, bookmarks, proofErr, permissions, comments, del, moveFrom…
                break;
        }
    }

    private void ReadRun(XmlReader r, ParagraphState state)
    {
        int depth = r.Depth;
        if (!r.Enter()) return;
        RunFormat? direct = null;
        RunProperties? props = null;
        RunProperties Props() => props ??= ResolveRun(state.StyleId, state.Cell, direct);

        while (r.NextChild(depth))
        {
            if (r.IsMC() && r.LocalName == "AlternateContent")
            {
                // Typically wraps a w:drawing; read the choice as run content.
                ReadAlternateContent(r, () =>
                {
                    int d = r.Depth;
                    if (!r.Enter()) return;
                    while (r.NextChild(d))
                    {
                        if (r.IsW() && r.LocalName == "drawing") ReadDrawing(r, state, Props());
                        else r.Skip();
                    }
                });
                continue;
            }
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "rPr":
                    direct = RunFormat.Parse(r);
                    props = null;
                    continue;
                case "t":
                    var text = r.ReadText();
                    if (!Suppressed(Props(), state, text)) AddText(state, text, Props());
                    continue;
                case "instrText":
                    var instruction = r.ReadText();
                    if (_fields.TryPeek(out var top) && !top.Separated) top.Instruction.Append(instruction);
                    continue;
                case "tab":
                case "ptab":
                    if (!Suppressed(Props(), state)) AddInline(state, new Tab(Props()));
                    break;
                case "br":
                case "cr":
                    if (!Suppressed(Props(), state))
                    {
                        var kind = r.WAttr("type") switch { "page" => BreakKind.Page, "column" => BreakKind.Column, _ => BreakKind.Line };
                        AddInline(state, new Break(r.LocalName == "cr" ? BreakKind.Line : kind, Props()));
                    }
                    break;
                case "noBreakHyphen":
                    if (!Suppressed(Props(), state, "‑")) AddText(state, "‑", Props());
                    break;
                case "sym":
                    if (!Suppressed(Props(), state)) AddSymbol(r, state, Props());
                    break;
                case "pgNum":
                    if (!Suppressed(Props(), state)) AddInline(state, new FieldRun(FieldKind.Page, Props()));
                    break;
                case "fldChar":
                    HandleFieldChar(r.WAttr("fldCharType"), state, Props());
                    break;
                case "drawing":
                    if (!Suppressed(Props(), state)) ReadDrawing(r, state, Props());
                    else r.Skip();
                    continue;
                case "pict":
                case "object":
                    context.WarnOnce("docx.vml", "VML pictures and embedded objects are not supported and were skipped.");
                    break;
                case "footnoteReference":
                case "endnoteReference":
                    context.WarnOnce("docx.notes", "Footnotes and endnotes are not supported yet; references were skipped.");
                    break;
            }
            r.Skip();
        }
    }

    private void AddText(ParagraphState state, string text, RunProperties props)
    {
        if (text.Length == 0) return;
        if (SymbolFonts.IsSymbolFont(props.FontFamily)) text = SymbolFonts.Map(props.FontFamily, text);
        text = text.Replace("­", ""); // soft hyphens are invisible unless a line breaks there
        if (text.Length > 0) AddInline(state, new TextRun(text, props));
    }

    /// <summary>Adds an inline to the paragraph, attaching the current hyperlink and recording bookmark text.</summary>
    private void AddInline(ParagraphState state, Inline inline)
    {
        inline.Link ??= CurrentLink();
        if (inline is TextRun text && _openBookmarks.Count > 0)
            foreach (var name in _openBookmarks.Values)
                if (context.Bookmarks.TryGetValue(name, out var sb) && sb.Length < 4096) sb.Append(text.Text);
        state.Paragraph.Inlines.Add(inline);
    }

    private Hyperlink? CurrentLink()
    {
        foreach (var f in _fields)
            if (f.Separated && f.Field?.Link is { } link) return link;
        return _links.TryPeek(out var current) ? current : null;
    }

    private Hyperlink? HyperlinkTarget(XmlReader r)
    {
        var anchor = r.WAttr("anchor");
        string? uri = null;
        if (r.RAttr("id") is { } id && context.Package.GetRelationships(partName).TryGetValue(id, out var rel) && rel.IsExternal)
            uri = rel.Target;
        return uri is null && anchor is null ? null : new Hyperlink(uri, anchor);
    }

    private void AddSymbol(XmlReader r, ParagraphState state, RunProperties props)
    {
        if (!int.TryParse(r.WAttr("char"), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code) || code is <= 0 or > 0xFFFF)
            return;
        var font = r.WAttr("font");
        var text = ((char)code).ToString();
        if (SymbolFonts.IsSymbolFont(font)) text = SymbolFonts.Map(font!, text, unknown: '□');
        else if (font is not null) props = props with { FontFamily = font };
        AddInline(state, new TextRun(text, props));
    }

    private static void MergeAdjacentText(List<Inline> inlines)
    {
        for (int i = inlines.Count - 1; i > 0; i--)
        {
            if (inlines[i] is TextRun b && inlines[i - 1] is TextRun a && ReferenceEquals(a.Properties, b.Properties) && a.Link == b.Link)
            {
                inlines[i - 1] = new TextRun(a.Text + b.Text, a.Properties) { Link = a.Link };
                inlines.RemoveAt(i);
            }
        }
    }

    // ---------------------------------------------------------------- fields

    /// <summary>
    /// Decides whether run content is shown. Instructions are always hidden. For fields we evaluate
    /// ourselves (PAGE, PAGEREF, REF…) the first piece of cached result is replaced by a <see cref="FieldRun"/>
    /// (or a REF placeholder) carrying that run's formatting, and the rest of the cached result is hidden.
    /// </summary>
    private bool Suppressed(RunProperties props, ParagraphState state, string? text = null)
    {
        if (_fields.Count == 0) return false;
        foreach (var f in _fields)
            if (!f.Separated) return true;
        foreach (var f in _fields)
        {
            if (f.Field is not { IsEvaluated: true } field)
            {
                f.HasResult = true;
                continue;
            }
            if (f.Emitted is null)
            {
                f.Emitted = field.CreateInline(props);
                AddInline(state, f.Emitted);
            }
            if (text is not null) AppendCached(f.Emitted, text);
            return true;
        }
        return false;
    }

    private static void AppendCached(Inline inline, string text)
    {
        switch (inline)
        {
            case FieldRun run: run.CachedResult += text; break;
            case RefPlaceholder placeholder: placeholder.CachedResult += text; break;
        }
    }

    private void HandleFieldChar(string? type, ParagraphState state, RunProperties props)
    {
        switch (type)
        {
            case "begin":
                _fields.Push(new FieldFrame { Properties = props });
                break;
            case "separate":
                if (_fields.TryPeek(out var frame)) frame.Separate();
                break;
            case "end":
                if (_fields.TryPeek(out var ending)) ending.Separate();
                EndField(state);
                break;
        }
    }

    /// <summary>Closes the innermost field; an evaluated field without cached result is emitted now.</summary>
    private void EndField(ParagraphState state)
    {
        if (!_fields.TryPop(out var frame)) return;
        bool hidden = _fields.Any(f => !f.Separated || f.Emitted is not null);
        if (hidden || frame.Field is not { } field) return;
        var props = frame.Properties ?? state.Paragraph.MarkProperties;
        if (field.IsEvaluated && frame.Emitted is null)
            AddInline(state, field.CreateInline(props));
        else if (!frame.HasResult && field.Name is "DATE" or "TIME")
            AddText(state, FieldInstruction.FormatDate(DateTime.Now, field.Picture ?? (field.Name == "TIME" ? "h:mm am/pm" : "M/d/yyyy")), props);
    }

    // ---------------------------------------------------------------- drawings

    private void ReadDrawing(XmlReader r, ParagraphState state, RunProperties props)
    {
        int depth = r.Depth;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (r.IsWP() && r.LocalName == "inline") ReadInlineDrawing(r, state, props);
            else
            {
                if (r.IsWP() && r.LocalName == "anchor")
                    context.WarnOnce("docx.floating", "Floating (anchored) drawings are not supported yet and were skipped.");
                r.Skip();
            }
        }
    }

    private void ReadInlineDrawing(XmlReader r, ParagraphState state, RunProperties props)
    {
        long cx = 0, cy = 0;
        string? embed = null;
        string? graphicType = null;
        int depth = r.Depth;
        if (r.IsEmptyElement) { r.Read(); return; }
        r.Read();
        while (!r.EOF && r.Depth > depth)
        {
            if (r.NodeType == XmlNodeType.Element)
            {
                if (r.IsWP() && r.LocalName == "extent")
                {
                    long.TryParse(r.GetAttribute("cx"), NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
                    long.TryParse(r.GetAttribute("cy"), NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
                }
                else if (r.IsA() && r.LocalName == "blip") embed ??= r.RAttr("embed");
                else if (r.IsA() && r.LocalName == "graphicData") graphicType = r.GetAttribute("uri");
            }
            r.Read();
        }
        if (r.NodeType == XmlNodeType.EndElement && r.Depth == depth) r.Read();

        if (embed is null)
        {
            var kind = graphicType?[(graphicType.LastIndexOf('/') + 1)..] ?? "unknown";
            context.WarnOnce("docx.graphic." + kind, $"Inline graphic of type '{kind}' (charts, shapes, diagrams) is not supported.");
            return;
        }
        double width = Units.FromEmu(cx), height = Units.FromEmu(cy);
        if (width <= 0 || height <= 0) return;
        double max = context.Options.MaxDimension;
        if (width > max || height > max)
        {
            double scale = max / Math.Max(width, height);
            (width, height) = (width * scale, height * scale);
            context.Warn("docx.image-size", "An image with excessive dimensions was scaled down.");
        }
        if (!context.Package.GetRelationships(partName).TryGetValue(embed, out var rel) || rel.IsExternal)
        {
            context.Warn("docx.image-link", $"Image relationship '{embed}' is missing or external.");
            return;
        }
        if (context.GetImage(rel.Target) is { } image)
            state.Paragraph.Inlines.Add(new ImageRun(image, width, height, props));
    }

    // ---------------------------------------------------------------- sections

    private SectionData ReadSection(XmlReader r)
    {
        var d = new SectionData();
        int depth = r.Depth;
        if (!r.Enter()) return d;
        while (r.NextChild(depth))
        {
            if (!r.IsW()) { r.Skip(); continue; }
            switch (r.LocalName)
            {
                case "pgSz":
                    d.Width = r.TwipsAttr("w");
                    d.Height = r.TwipsAttr("h");
                    d.Landscape = r.WAttr("orient") == "landscape";
                    break;
                case "pgMar":
                    d.Top = r.TwipsAttr("top");
                    d.Right = r.TwipsAttr("right");
                    d.Bottom = r.TwipsAttr("bottom");
                    d.Left = r.TwipsAttr("left");
                    d.Header = r.TwipsAttr("header");
                    d.Footer = r.TwipsAttr("footer");
                    break;
                case "titlePg": d.TitlePage = r.OnOff(); break;
                case "type":
                    d.Start = r.WAttr("val") switch
                    {
                        "continuous" => SectionStart.Continuous,
                        "evenPage" => SectionStart.EvenPage,
                        "oddPage" => SectionStart.OddPage,
                        _ => SectionStart.NextPage,
                    };
                    break;
                case "pgNumType":
                    d.PageNumberStart = r.IntAttr("start");
                    d.PageNumberFormat = r.WAttr("fmt") switch
                    {
                        "upperRoman" => PageNumberFormat.UpperRoman,
                        "lowerRoman" => PageNumberFormat.LowerRoman,
                        "upperLetter" => PageNumberFormat.UpperLetter,
                        "lowerLetter" => PageNumberFormat.LowerLetter,
                        _ => PageNumberFormat.Decimal,
                    };
                    break;
                case "cols":
                    d.Columns = ReadColumns(r);
                    continue;
                case "headerReference":
                    if (r.RAttr("id") is { } hid) d.Headers[r.WAttr("type") ?? "default"] = hid;
                    break;
                case "footerReference":
                    if (r.RAttr("id") is { } fid) d.Footers[r.WAttr("type") ?? "default"] = fid;
                    break;
            }
            r.Skip();
        }
        return d;
    }

    private static SectionColumns ReadColumns(XmlReader r)
    {
        int count = Math.Clamp(r.IntAttr("num") ?? 1, 1, 45);
        double spacing = Math.Clamp(r.TwipsAttr("space") ?? 36, 0, 1584);
        bool separator = r.WAttr("sep") is not null && r.OnOff("sep");
        bool equal = r.OnOff("equalWidth");
        var definitions = new List<ColumnDefinition>();
        int depth = r.Depth;
        if (r.Enter())
        {
            while (r.NextChild(depth))
            {
                if (r.IsW() && r.LocalName == "col" && definitions.Count < 45)
                    definitions.Add(new ColumnDefinition(Math.Clamp(r.TwipsAttr("w") ?? 0, 0, 31_680), Math.Clamp(r.TwipsAttr("space") ?? 0, 0, 1584)));
                r.Skip();
            }
        }
        // Explicit widths only apply when equalWidth is off; w:num still gives the count otherwise.
        if (equal || definitions.Count < 2 || definitions.All(d => d.Width <= 0)) definitions.Clear();
        return new SectionColumns { Count = definitions.Count > 0 ? definitions.Count : count, Spacing = spacing, Definitions = definitions, Separator = separator };
    }

    // ---------------------------------------------------------------- tables

    private Table? ReadTable(XmlReader r, int nesting)
    {
        if (nesting > context.Options.MaxTableNesting)
        {
            context.WarnOnce("docx.table-nesting", "Tables nested too deeply were skipped.");
            r.Skip();
            return null;
        }
        // Table-style conditional formatting (lastRow…) needs the row count before cells are resolved,
        // so the table is buffered and parsed twice. Namespace declarations are preserved by ReadOuterXml.
        var xml = r.ReadOuterXml();
        int rowCount = CountRows(xml);
        using var tr = XmlReader.Create(new StringReader(xml), WordXml.Settings(xml.Length + 1));
        tr.MoveToContent();
        return ReadTableBuffered(tr, rowCount, nesting);
    }

    private static int CountRows(string xml)
    {
        using var r = XmlReader.Create(new StringReader(xml), WordXml.Settings(xml.Length + 1));
        var tables = new Stack<int>();
        int rows = 0;
        while (r.Read())
        {
            if (r.NodeType != XmlNodeType.Element || !r.IsW()) continue;
            while (tables.Count > 0 && tables.Peek() >= r.Depth) tables.Pop();
            if (r.LocalName == "tbl" && !r.IsEmptyElement) tables.Push(r.Depth);
            else if (r.LocalName == "tr" && tables.Count == 1) rows++;
        }
        return rows;
    }

    private sealed class RowBuilder
    {
        public required RowFormat Format { get; init; }
        public List<(CellFormat Format, string? Xml)> Cells { get; } = [];
    }

    private Table ReadTableBuffered(XmlReader r, int rowCount, int nesting)
    {
        var direct = new TableFormat();
        var grid = new List<double>();
        var rows = new List<RowBuilder>();
        int depth = r.Depth;
        if (r.Enter())
        {
            while (r.NextChild(depth))
            {
                if (!r.IsW()) { r.Skip(); continue; }
                switch (r.LocalName)
                {
                    case "tblPr": direct = TableFormat.Parse(r); break;
                    case "tblGrid": ReadGrid(r, grid); break;
                    case "tr": rows.Add(ReadRow(r)); break;
                    case "sdt": ReadContainer(r, "sdtContent", () => ReadRowsIn(r, rows)); break;
                    case "customXml": ReadRowsIn(r, rows); break;
                    default: r.Skip(); break;
                }
            }
        }

        var styleId = direct.StyleId is { } id && Styles.Get(id)?.Type == StyleType.Table ? id : Styles.DefaultTableStyle;
        var style = Styles.Resolve(styleId);
        var format = style.Table.Merge(direct);
        int look = direct.Look ?? style.Table.Look ?? (TableFormat.LookFirstRow | TableFormat.LookFirstColumn | TableFormat.LookNoVBand);
        int columns = Math.Max(grid.Count, rows.Count == 0 ? 0 : rows.Max(row => (row.Format.GridBefore ?? 0) + row.Cells.Sum(c => c.Format.GridSpan ?? 1)));

        var table = new Table
        {
            Properties = new TableProperties
            {
                Width = format.Width ?? TableWidth.Auto,
                Alignment = format.Alignment ?? TableAlignment.Left,
                Indent = format.Indent ?? 0,
                Layout = format.Layout ?? TableLayoutMode.Auto,
            },
        };
        table.Grid.AddRange(grid);

        for (int ri = 0; ri < rows.Count; ri++)
        {
            var raw = rows[ri];
            var row = new TableRow
            {
                Height = Math.Clamp(raw.Format.Height ?? 0, 0, context.Options.MaxDimension),
                HeightRule = raw.Format.Height is null ? RowHeightRule.Auto : raw.Format.HeightRule ?? RowHeightRule.AtLeast,
                IsHeader = raw.Format.IsHeader ?? false,
                CantSplit = raw.Format.CantSplit ?? false,
                GridBefore = raw.Format.GridBefore ?? 0,
            };
            int column = row.GridBefore;
            foreach (var (cellFormat, cellXml) in raw.Cells)
            {
                int span = cellFormat.GridSpan ?? 1;
                var conditional = Conditional(style, look, format, ri, rowCount, column, span, columns);
                var cellStyle = new CellStyle(style.Para.Merge(conditional.Para), style.Run.Merge(conditional.Run));
                var effective = style.Cell.Merge(conditional.Cell);
                var cell = new TableCell
                {
                    Width = cellFormat.Width ?? TableWidth.Auto,
                    GridSpan = span,
                    VerticalMerge = cellFormat.VerticalMerge ?? VerticalMerge.None,
                    Borders = CellBorders(format.Borders, effective.Borders, cellFormat.Borders, ri == 0, ri == rowCount - 1,
                        column == 0, column + span >= columns),
                    Shading = (cellFormat.Shading ?? effective.Shading ?? format.Shading)?.Value,
                    Margins = CellMargins(format.CellMargins.Merge(effective.Margins).Merge(cellFormat.Margins)),
                    VerticalAlignment = cellFormat.VerticalAlignment ?? effective.VerticalAlignment ?? CellVerticalAlignment.Top,
                };
                if (cellXml is not null)
                {
                    using var cr = XmlReader.Create(new StringReader(cellXml), WordXml.Settings(cellXml.Length + 1));
                    cr.MoveToContent();
                    ReadBlocks(cr, new BlockSink(cell.Blocks, null), cellStyle, nesting);
                }
                row.Cells.Add(cell);
                column += span;
            }
            table.Rows.Add(row);
        }
        return table;
    }

    private void ReadRowsIn(XmlReader r, List<RowBuilder> rows)
    {
        int depth = r.Depth;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (r.IsW() && r.LocalName == "tr") rows.Add(ReadRow(r));
            else r.Skip();
        }
    }

    private static void ReadGrid(XmlReader r, List<double> grid)
    {
        int depth = r.Depth;
        if (!r.Enter()) return;
        while (r.NextChild(depth))
        {
            if (r.IsW() && r.LocalName == "gridCol" && grid.Count < 64)
                grid.Add(Math.Clamp(r.TwipsAttr("w") ?? 0, 0, 14_400));
            r.Skip();
        }
    }

    private RowBuilder ReadRow(XmlReader r)
    {
        var format = new RowFormat();
        var cells = new List<(CellFormat, string?)>();
        int depth = r.Depth;
        if (r.Enter())
        {
            while (r.NextChild(depth))
            {
                if (r.IsW() && r.LocalName == "trPr") format = RowFormat.Parse(r);
                else if (r.IsW() && r.LocalName == "tc") ReadCell(r, cells);
                else if (r.IsW() && r.LocalName is "sdt" or "customXml")
                {
                    void ReadCells()
                    {
                        int d = r.Depth;
                        if (!r.Enter()) return;
                        while (r.NextChild(d))
                        {
                            if (r.IsW() && r.LocalName == "tc") ReadCell(r, cells);
                            else r.Skip();
                        }
                    }
                    if (r.LocalName == "sdt") ReadContainer(r, "sdtContent", ReadCells); else ReadCells();
                }
                else r.Skip();
            }
        }
        var row = new RowBuilder { Format = format };
        row.Cells.AddRange(cells);
        return row;
    }

    /// <summary>Reads tcPr now and keeps the cell body as XML until the table-style context is known.</summary>
    private static void ReadCell(XmlReader r, List<(CellFormat, string?)> cells)
    {
        var xml = r.ReadOuterXml();
        using var cr = XmlReader.Create(new StringReader(xml), WordXml.Settings(xml.Length + 1));
        cr.MoveToContent();
        var format = new CellFormat();
        int depth = cr.Depth;
        if (cr.Enter())
        {
            while (cr.NextChild(depth))
            {
                if (cr.IsW() && cr.LocalName == "tcPr") format = CellFormat.Parse(cr);
                else cr.Skip();
            }
        }
        // Legacy horizontal merge: fold continuation cells into the previous cell's span.
        if (format.HorizontalMergeContinue == true && cells.Count > 0)
        {
            var (prev, prevXml) = cells[^1];
            cells[^1] = (prev with { GridSpan = (prev.GridSpan ?? 1) + (format.GridSpan ?? 1) }, prevXml);
            return;
        }
        cells.Add((format, xml));
    }

    private static ConditionalFormat Conditional(ResolvedStyle style, int look, TableFormat table, int row, int rowCount,
        int column, int span, int columns)
    {
        bool firstRow = (look & TableFormat.LookFirstRow) != 0, lastRow = (look & TableFormat.LookLastRow) != 0;
        bool firstCol = (look & TableFormat.LookFirstColumn) != 0, lastCol = (look & TableFormat.LookLastColumn) != 0;
        bool isFirstRow = firstRow && row == 0, isLastRow = lastRow && row == rowCount - 1 && rowCount > 1;
        bool isFirstCol = firstCol && column == 0, isLastCol = lastCol && column + span >= columns;

        var kinds = new List<string> { "wholeTable" };
        if ((look & TableFormat.LookNoVBand) == 0 && !isFirstCol && !isLastCol)
        {
            int band = (column - (firstCol ? 1 : 0)) / Math.Max(1, table.ColumnBandSize ?? 1);
            kinds.Add(band % 2 == 0 ? "band1Vert" : "band2Vert");
        }
        if ((look & TableFormat.LookNoHBand) == 0 && !isFirstRow && !isLastRow)
        {
            int band = (row - (firstRow ? 1 : 0)) / Math.Max(1, table.RowBandSize ?? 1);
            kinds.Add(band % 2 == 0 ? "band1Horz" : "band2Horz");
        }
        if (isFirstCol) kinds.Add("firstCol");
        if (isLastCol) kinds.Add("lastCol");
        if (isFirstRow) kinds.Add("firstRow");
        if (isLastRow) kinds.Add("lastRow");
        if (isFirstRow && isFirstCol) kinds.Add("nwCell");
        if (isFirstRow && isLastCol) kinds.Add("neCell");
        if (isLastRow && isFirstCol) kinds.Add("swCell");
        if (isLastRow && isLastCol) kinds.Add("seCell");

        var result = new ConditionalFormat(null, null, null, null);
        foreach (var kind in kinds)
            if (style.Conditionals.TryGetValue(kind, out var c)) result = result.Merge(c);
        return result;
    }

    private static BorderSet CellBorders(EdgeBorders table, EdgeBorders conditional, EdgeBorders cell,
        bool top, bool bottom, bool left, bool right)
    {
        Border Edge(Border? own, Border? cond, Border? outer, Border? condInside, Border? inside, bool isOuter) =>
            own ?? cond ?? (isOuter ? outer : condInside ?? inside) ?? Border.None;

        return new BorderSet
        {
            Top = Edge(cell.Top, conditional.Top, table.Top, conditional.InsideH, table.InsideH, top),
            Bottom = Edge(cell.Bottom, conditional.Bottom, table.Bottom, conditional.InsideH, table.InsideH, bottom),
            Left = Edge(cell.Left, conditional.Left, table.Left, conditional.InsideV, table.InsideV, left),
            Right = Edge(cell.Right, conditional.Right, table.Right, conditional.InsideV, table.InsideV, right),
        };
    }

    private static Thickness CellMargins(EdgeLengths m) =>
        new(Math.Clamp(m.Top ?? 0, 0, 720), Math.Clamp(m.Right ?? 5.4, 0, 720), Math.Clamp(m.Bottom ?? 0, 0, 720), Math.Clamp(m.Left ?? 5.4, 0, 720));
}