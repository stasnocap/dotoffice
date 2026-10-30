using System.Xml;
using DotOffice.Core;

namespace DotOffice.Docx;

/// <summary>Reads WordprocessingML (.docx) packages into the format-neutral <see cref="Document"/> model.</summary>
public static class DocxReader
{
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

        using var package = new OpcPackage(stream, options);
        var main = package.FindRelationship("", "officeDocument")
                   ?? throw new InvalidDataException("Package has no main document (officeDocument relationship).");
        if (!package.Exists(main.Target))
            throw new InvalidDataException($"Main document part '{main.Target}' is missing.");

        var context = new DocxContext(package, options, main.Target);
        return context.ReadDocument();
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

internal sealed record DocxSettings(double DefaultTabStop, bool EvenAndOddHeaders, int CompatibilityMode, bool DoNotExpandShiftReturn)
{
    public static DocxSettings Default { get; } = new(36, false, 15, false);

    public static DocxSettings Read(XmlReader r)
    {
        var s = Default with { CompatibilityMode = 12 };
        while (r.Read())
        {
            if (r.NodeType != XmlNodeType.Element || !r.IsW()) continue;
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
    public int Columns = 1;
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
        using (var r = Package.OpenXml(MainPart))
        {
            r.MoveToContent();
            while (r.Read())
            {
                if (r.NodeType != XmlNodeType.Element || !r.IsW() || r.LocalName != "body") continue;
                story.ReadBlocks(r, blocks, cell: null, nesting: 0, onSection: data =>
                {
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
            var section = BuildSection(new SectionData(), previous, blocks);
            document.Sections.Add(section);
        }
        return document;
    }

    private Section BuildSection(SectionData d, Section? previous, List<Block> blocks)
    {
        double max = Options.MaxDimension;
        double width = Math.Clamp(d.Width ?? 612, 36, max), height = Math.Clamp(d.Height ?? 792, 36, max);
        if (d.Columns > 1)
            WarnOnce("docx.columns", "Multi-column sections are not supported yet; text is laid out in a single column.");

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