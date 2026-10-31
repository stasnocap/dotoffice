using System.Text;
using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Reads WordprocessingML (.docx) packages into the format-neutral <see cref="Document"/> model.</summary>
public static class DocxReader
{
    /// <summary>Reads a DOCX package. Malformed or oversized input throws <see cref="InvalidDataException"/>.</summary>
    public static Document Read(Stream stream, DocxReaderOptions? options = null)
    {
        options ??= new DocxReaderOptions();
        if (!stream.CanSeek)
        {
            var buffer = new MemoryStream();
            CopyLimited(stream, buffer, options.MaxTotalUncompressedSize);
            buffer.Position = 0;
            stream = buffer;
        }

        try
        {
            using var package = new OpcPackage(stream, options);
            var main = package.FindRelationship("", "officeDocument")
                       ?? throw new InvalidDataException("Package has no main document (officeDocument relationship).");
            if (!package.Exists(main.Target))
                throw new InvalidDataException($"Main document part '{main.Target}' is missing.");
            return new DocxContext(package, options, main.Target).ReadDocument();
        }
        catch (XmlException e)
        {
            throw new InvalidDataException($"The document contains malformed XML: {e.Message}", e);
        }
    }

    private static void CopyLimited(Stream source, Stream target, long limit)
    {
        var buffer = new byte[81920];
        long total = 0;
        int n;
        while ((n = source.Read(buffer)) > 0)
        {
            total += n;
            if (total > limit) throw new InvalidDataException($"Input is larger than the limit of {limit} bytes.");
            target.Write(buffer, 0, n);
        }
    }
}

/// <summary>Numbering of footnotes or endnotes (w:footnotePr/w:endnotePr).</summary>
internal sealed record NoteNumbering(string Format, int Start, string Restart = "continuous");

internal sealed record DocxSettings(double DefaultTabStop, bool EvenAndOddHeaders, int CompatibilityMode, bool DoNotExpandShiftReturn)
{
    public static DocxSettings Default { get; } = new(36, false, 15, false);

    public NoteNumbering Footnotes { get; init; } = new("decimal", 1);
    public NoteNumbering Endnotes { get; init; } = new("lowerRoman", 1);

    public static DocxSettings Read(XmlReader r)
    {
        var s = Default with { CompatibilityMode = 12 };
        string? notes = null; // "footnotePr" or "endnotePr" while inside one
        int notesDepth = 0;
        while (r.Read())
        {
            if (r.NodeType != XmlNodeType.Element || !r.IsW()) continue;
            if (notes is not null && r.Depth <= notesDepth) notes = null;
            switch (r.LocalName)
            {
                case "defaultTabStop":
                    if (r.TwipsAttr("val") is { } tab && tab > 1) s = s with { DefaultTabStop = Math.Min(tab, 720) };
                    break;
                case "evenAndOddHeaders": s = s with { EvenAndOddHeaders = r.OnOff() }; break;
                case "doNotExpandShiftReturn": s = s with { DoNotExpandShiftReturn = r.OnOff() }; break;
                case "compatSetting" when r.WAttr("name") == "compatibilityMode":
                    if (int.TryParse(r.WAttr("val"), out var mode)) s = s with { CompatibilityMode = mode };
                    break;
                case "footnotePr" or "endnotePr":
                    (notes, notesDepth) = (r.LocalName, r.Depth);
                    break;
                case "numFmt" or "numStart" or "numRestart" when notes is not null:
                    var numbering = notes == "footnotePr" ? s.Footnotes : s.Endnotes;
                    numbering = r.LocalName switch
                    {
                        "numFmt" => numbering with { Format = r.WAttr("val") ?? numbering.Format },
                        "numStart" => numbering with { Start = Math.Clamp(r.IntAttr("val") ?? 1, 0, 100_000) },
                        _ => numbering with { Restart = r.WAttr("val") ?? numbering.Restart },
                    };
                    s = notes == "footnotePr" ? s with { Footnotes = numbering } : s with { Endnotes = numbering };
                    break;
            }
        }
        return s;
    }
}

/// <summary>Raw w:sectPr values before inheritance and defaults are applied.</summary>
internal sealed class SectionData
{
    public double? Width, Height, Top, Right, Bottom, Left, Header, Footer;
    public bool Landscape, TitlePage;
    public SectionStart Start;
    public int? PageNumberStart;
    public PageNumberFormat PageNumberFormat;
    public SectionColumns Columns = SectionColumns.Single;

    /// <summary>w:numRestart of the section's w:footnotePr/w:endnotePr, overriding the document settings.</summary>
    public string? FootnoteRestart, EndnoteRestart;
    public Dictionary<string, string> Headers { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, string> Footers { get; } = new(StringComparer.Ordinal);
}

/// <summary>Package-wide state: parts shared by all stories, caches and limits.</summary>
internal sealed class DocxContext
{
    private readonly Dictionary<string, HeaderFooter> _headerFooters = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ImageData?> _images = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _warned = new(StringComparer.Ordinal);

    public DocxContext(OpcPackage package, DocxReaderOptions options, string mainPart)
    {
        Package = package;
        Options = options;
        MainPart = mainPart;

        var theme = ReadPart(package.FindRelationship(mainPart, "theme"), Theme.Read);
        Styles = ReadPart(package.FindRelationship(mainPart, "styles"), StyleSheet.Read) ?? new StyleSheet();
        Styles.Theme = theme;
        Numbering = ReadPart(package.FindRelationship(mainPart, "numbering"), Numbering.Read) ?? new Numbering();
        Settings = ReadPart(package.FindRelationship(mainPart, "settings"), DocxSettings.Read) ?? DocxSettings.Default;
    }

    public OpcPackage Package { get; }
    public DocxReaderOptions Options { get; }
    public string MainPart { get; }
    public StyleSheet Styles { get; }
    public Numbering Numbering { get; }
    public DocxSettings Settings { get; }

    private readonly Dictionary<NoteKind, Dictionary<string, string>?> _noteParts = [];
    private readonly List<Note> _notes = [];
    private int _footnoteCount, _endnoteCount, _noteDepth;

    /// <summary>Endnotes in reference order.</summary>
    public List<Note> Endnotes { get; } = [];

    /// <summary>
    /// Reads the note <paramref name="id"/> for a new reference and numbers it. Returns null when the
    /// note does not exist (or notes nest too deeply).
    /// </summary>
    public Note? ReadNote(NoteKind kind, string id, bool customMark)
    {
        var part = kind == NoteKind.Footnote ? "footnotes" : "endnotes";
        if (!_noteParts.TryGetValue(kind, out var notes)) _noteParts[kind] = notes = LoadNotes(part);
        if (notes is null || !notes.TryGetValue(id, out var xml) || _noteDepth > 2)
        {
            Warn("docx.note-missing", $"The {kind.ToString().ToLowerInvariant()} with id '{id}' was not found.");
            return null;
        }

        var numbering = kind == NoteKind.Footnote ? Settings.Footnotes : Settings.Endnotes;
        int number = numbering.Start + (kind == NoteKind.Footnote ? _footnoteCount++ : _endnoteCount++);
        var mark = customMark ? "" : Numbering.FormatNumber(number, numbering.Format);
        var note = new Note(kind)
        {
            Mark = mark,
            PageMark = kind == NoteKind.Footnote && numbering.Restart == "eachPage" && !customMark
                ? n => Numbering.FormatNumber(numbering.Start + n - 1, numbering.Format)
                : null,
        };
        _notes.Add(note);
        if (kind == NoteKind.Endnote) Endnotes.Add(note);

        _noteDepth++;
        try
        {
            var partName = Package.FindRelationship(MainPart, part)!.Target;
            using var r = XmlReader.Create(new StringReader(xml), WordXml.Settings(xml.Length + 1));
            r.MoveToContent();
            new StoryReader(this, partName) { NoteOwner = note }.ReadBlocks(r, note.Blocks, cell: null, nesting: 0, onSection: null);
        }
        finally
        {
            _noteDepth--;
        }
        return note;
    }

    /// <summary>Note bodies by id; separator notes are skipped because layout draws its own separator.</summary>
    private Dictionary<string, string>? LoadNotes(string part)
    {
        var rel = Package.FindRelationship(MainPart, part);
        if (rel is null || !Package.Exists(rel.Target)) return null;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        try
        {
            using var r = Package.OpenXml(rel.Target);
            r.MoveToContent();
            while (!r.EOF)
            {
                if (r.NodeType == XmlNodeType.Element && r.IsW() && r.LocalName is "footnote" or "endnote")
                {
                    var id = r.WAttr("id");
                    var type = r.WAttr("type");
                    var xml = r.ReadOuterXml();
                    if (id is not null && type is null or "normal") result.TryAdd(id, xml);
                }
                else r.Read();
            }
        }
        catch (XmlException e)
        {
            Warn("docx.malformed-part", $"Ignoring malformed part '{rel.Target}': {e.Message}");
            return null;
        }
        return result;
    }

    /// <summary>Bookmark names and the text they enclose (for REF fields), shared by all stories.</summary>
    public Dictionary<string, StringBuilder> Bookmarks { get; } = new(StringComparer.OrdinalIgnoreCase);

    public void Warn(string code, string message) => Options.Warnings.Warn(code, message);

    /// <summary>Emits a warning once per code (for content that is skipped wholesale).</summary>
    public void WarnOnce(string code, string message)
    {
        if (_warned.Add(code)) Warn(code, message);
    }

    public Document ReadDocument()
    {
        var document = new Document
        {
            Settings = new DocumentSettings
            {
                DefaultTabStop = Settings.DefaultTabStop,
                EvenAndOddHeaders = Settings.EvenAndOddHeaders,
                TableEdgeAtIndent = Settings.CompatibilityMode >= 15,
                JustifyLinesEndingWithBreak = !Settings.DoNotExpandShiftReturn,
            },
        };
        ReadCoreProperties(document);

        var story = new StoryReader(this, MainPart);
        var blocks = new List<Block>();
        Section? previous = null;
        var sectionData = new List<SectionData>();
        using (var r = Package.OpenXml(MainPart))
        {
            r.MoveToContent();
            while (r.Read())
            {
                if (r.NodeType != XmlNodeType.Element || !r.IsW() || r.LocalName != "body") continue;
                story.ReadBlocks(r, blocks, cell: null, nesting: 0, onSection: data =>
                {
                    sectionData.Add(data);
                    previous = BuildSection(data, previous, blocks);
                    document.Sections.Add(previous);
                    blocks = [];
                    return blocks;
                });
                break;
            }
        }
        if (blocks.Count > 0 || document.Sections.Count == 0)
        {
            // No trailing sectPr: the remaining content uses the previous (or default) section settings.
            sectionData.Add(new SectionData());
            var section = BuildSection(sectionData[^1], previous, blocks);
            document.Sections.Add(section);
        }
        document.Endnotes.AddRange(Endnotes);
        ResolveReferences(document);
        RestartNoteNumbers(document, sectionData);
        return document;
    }

    /// <summary>Renumbers notes whose numbering restarts in each section (notes are numbered continuously while reading).</summary>
    private void RestartNoteNumbers(Document document, List<SectionData> sections)
    {
        foreach (var kind in new[] { NoteKind.Footnote, NoteKind.Endnote })
        {
            var numbering = kind == NoteKind.Footnote ? Settings.Footnotes : Settings.Endnotes;
            string Restart(int i) => (kind == NoteKind.Footnote ? sections[i].FootnoteRestart : sections[i].EndnoteRestart) ?? numbering.Restart;
            if (!Enumerable.Range(0, sections.Count).Any(i => Restart(i) == "eachSect")) continue;

            int count = 0;
            for (int i = 0; i < document.Sections.Count; i++)
            {
                if (i == 0 || Restart(i) == "eachSect") count = 0;
                foreach (var reference in Paragraphs(document.Sections[i].Blocks).SelectMany(p => p.Inlines).OfType<NoteReference>())
                    if (reference.Note.Kind == kind && reference.Mark.Length > 0)
                        reference.Note.Mark = Numbering.FormatNumber(numbering.Start + count++, numbering.Format);
            }
        }
    }

    /// <summary>Replaces REF placeholders by the bookmarked text (forward references are allowed).</summary>
    private void ResolveReferences(Document document)
    {
        var stories = document.Sections.Select(s => s.Blocks)
            .Concat(_headerFooters.Values.Select(h => h.Blocks))
            .Concat(_notes.Select(n => n.Blocks));
        foreach (var blocks in stories)
            foreach (var paragraph in Paragraphs(blocks))
            {
                var inlines = paragraph.Inlines;
                for (int i = 0; i < inlines.Count; i++)
                {
                    if (inlines[i] is not RefPlaceholder placeholder) continue;
                    var text = Bookmarks.TryGetValue(placeholder.Bookmark, out var sb) && sb.Length > 0
                        ? sb.ToString()
                        : placeholder.CachedResult;
                    if (string.IsNullOrEmpty(text))
                    {
                        Warn("docx.ref-missing", $"REF field points to unknown bookmark '{placeholder.Bookmark}'.");
                        text = "Error! Reference source not found.";
                    }
                    inlines[i] = new TextRun(text, placeholder.Properties) { Link = placeholder.Link };
                }
            }
    }

    private static IEnumerable<Paragraph> Paragraphs(IEnumerable<Block> blocks)
    {
        foreach (var block in blocks)
        {
            if (block is Paragraph p) yield return p;
            else if (block is Table t)
                foreach (var cell in t.Rows.SelectMany(r => r.Cells))
                    foreach (var nested in Paragraphs(cell.Blocks))
                        yield return nested;
        }
    }

    private Section BuildSection(SectionData d, Section? previous, List<Block> blocks)
    {
        double max = Options.MaxDimension;
        double width = Math.Clamp(d.Width ?? 612, 36, max), height = Math.Clamp(d.Height ?? 792, 36, max);

        var section = new Section
        {
            PageWidth = width,
            PageHeight = height,
            Orientation = d.Landscape || width > height ? Orientation.Landscape : Orientation.Portrait,
            Margins = new Thickness(Clamp(d.Top, 72, height), Clamp(d.Right, 72, width), Clamp(d.Bottom, 72, height), Clamp(d.Left, 72, width)),
            HeaderDistance = Clamp(d.Header, 36, height),
            FooterDistance = Clamp(d.Footer, 36, height),
            Start = d.Start,
            DifferentFirstPage = d.TitlePage,
            PageNumberStart = d.PageNumberStart,
            PageNumberFormat = d.PageNumberFormat,
            Columns = d.Columns,
            DefaultHeader = HeaderFooterFor(d.Headers, "default") ?? previous?.DefaultHeader,
            FirstHeader = HeaderFooterFor(d.Headers, "first") ?? previous?.FirstHeader,
            EvenHeader = HeaderFooterFor(d.Headers, "even") ?? previous?.EvenHeader,
            DefaultFooter = HeaderFooterFor(d.Footers, "default") ?? previous?.DefaultFooter,
            FirstFooter = HeaderFooterFor(d.Footers, "first") ?? previous?.FirstFooter,
            EvenFooter = HeaderFooterFor(d.Footers, "even") ?? previous?.EvenFooter,
        };
        section.Blocks.AddRange(blocks);
        return section;

        static double Clamp(double? value, double fallback, double limit) =>
            Math.Clamp(Math.Abs(value ?? fallback), 0, limit / 2);
    }

    private HeaderFooter? HeaderFooterFor(Dictionary<string, string> refs, string type)
    {
        if (!refs.TryGetValue(type, out var relId)) return null;
        if (!Package.GetRelationships(MainPart).TryGetValue(relId, out var rel) || rel.IsExternal || !Package.Exists(rel.Target))
        {
            Warn("docx.header-missing", $"Header/footer relationship '{relId}' does not resolve to a part.");
            return null;
        }
        if (_headerFooters.TryGetValue(rel.Target, out var cached)) return cached;

        var result = new HeaderFooter();
        _headerFooters[rel.Target] = result; // registered first so self-references cannot recurse
        var story = new StoryReader(this, rel.Target);
        using var r = Package.OpenXml(rel.Target);
        r.MoveToContent();
        if (r.NodeType == XmlNodeType.Element)
            story.ReadBlocks(r, result.Blocks, cell: null, nesting: 0, onSection: null);
        return result;
    }

    /// <summary>Loads an image part once; returns null (with a warning) if it is missing or too large.</summary>
    public ImageData? GetImage(string partName)
    {
        if (_images.TryGetValue(partName, out var cached)) return cached;
        ImageData? image = null;
        if (Package.Exists(partName))
        {
            try { image = new ImageData(Package.ReadAllBytes(partName), partName); }
            catch (InvalidDataException e) { Warn("docx.image-unreadable", $"Image '{partName}' skipped: {e.Message}"); }
        }
        else Warn("docx.image-missing", $"Image part '{partName}' is missing.");
        _images[partName] = image;
        return image;
    }

    private void ReadCoreProperties(Document document)
    {
        var core = Package.GetRelationships("").Values.FirstOrDefault(r => r.Kind == "core-properties" && !r.IsExternal);
        if (core is null || !Package.Exists(core.Target)) return;
        try
        {
            using var r = Package.OpenXml(core.Target);
            while (r.Read())
            {
                if (r.NodeType != XmlNodeType.Element) continue;
                if (r.LocalName == "title" && document.Title is null) document.Title = NonEmpty(r.ReadText());
                else if (r.LocalName == "creator" && document.Author is null) document.Author = NonEmpty(r.ReadText());
            }
        }
        catch (XmlException e)
        {
            Warn("docx.core-properties", $"Ignoring malformed document properties: {e.Message}");
        }

        static string? NonEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();
    }

    private T? ReadPart<T>(Relationship? rel, Func<XmlReader, T> read) where T : class
    {
        if (rel is null || !Package.Exists(rel.Target)) return null;
        try
        {
            using var r = Package.OpenXml(rel.Target);
            return read(r);
        }
        catch (XmlException e)
        {
            Warn("docx.malformed-part", $"Ignoring malformed part '{rel.Target}': {e.Message}");
            return null;
        }
    }
}