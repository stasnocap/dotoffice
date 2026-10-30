using System.IO.Compression;
using System.Text;

namespace DotOffice.Tests.Docx;

/// <summary>
/// Test-only DOCX writer: callers provide raw WordprocessingML fragments, the builder adds the
/// package plumbing (content types, relationships, namespace declarations).
/// </summary>
internal sealed class DocxBuilder
{
    public const string Namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" mc:Ignorable=\"w14\"";

    private readonly List<(string Id, string Type, string Target)> _relationships = [];
    private readonly Dictionary<string, byte[]> _parts = [];
    private int _nextId = 1;

    public string Body { get; set; } = "";
    public string? Styles { get; set; }
    public string? Numbering { get; set; }
    public string? Settings { get; set; }
    public string? Theme { get; set; }
    public string? Title { get; set; }

    /// <summary>Adds a header part and returns its relationship id.</summary>
    public string AddHeader(string innerXml) => AddStory("header", "hdr", innerXml);

    public string AddFooter(string innerXml) => AddStory("footer", "ftr", innerXml);

    /// <summary>Adds a media part and returns its relationship id.</summary>
    public string AddImage(string fileName, byte[] data)
    {
        var id = NextId();
        _parts[$"word/media/{fileName}"] = data;
        _relationships.Add((id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", $"media/{fileName}"));
        return id;
    }

    /// <summary>Inline picture markup (w:drawing) for an image relationship; size in points.</summary>
    public static string InlineImage(string relId, double widthPt, double heightPt) =>
        $"<w:r><w:drawing><wp:inline><wp:extent cx=\"{(long)(widthPt * 12700)}\" cy=\"{(long)(heightPt * 12700)}\"/>" +
        "<wp:docPr id=\"1\" name=\"Picture\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        $"<pic:pic><pic:blipFill><a:blip r:embed=\"{relId}\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    public byte[] Build()
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var rels = new List<(string, string, string)>(_relationships);
            void Optional(string? xml, string root, string file, string type)
            {
                if (xml is null) return;
                Write(zip, $"word/{file}", $"<w:{root} {Namespaces}>{xml}</w:{root}>");
                rels.Add((NextId(), $"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{type}", file));
            }
            Optional(Styles, "styles", "styles.xml", "styles");
            Optional(Numbering, "numbering", "numbering.xml", "numbering");
            Optional(Settings, "settings", "settings.xml", "settings");
            if (Theme is not null)
            {
                Write(zip, "word/theme/theme1.xml",
                    $"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"><a:themeElements>{Theme}</a:themeElements></a:theme>");
                rels.Add((NextId(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"));
            }
            foreach (var (name, data) in _parts)
            {
                using var s = zip.CreateEntry(name).Open();
                s.Write(data);
            }

            Write(zip, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/>" +
                "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            var packageRels = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>";
            if (Title is not null)
            {
                packageRels += "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
                Write(zip, "docProps/core.xml",
                    "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                    $"xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>{Title}</dc:title><dc:creator>Tests</dc:creator></cp:coreProperties>");
            }
            Write(zip, "_rels/.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{packageRels}</Relationships>");
            Write(zip, "word/document.xml", $"<w:document {Namespaces}><w:body>{Body}</w:body></w:document>");
            Write(zip, "word/_rels/document.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                string.Concat(rels.Select(r => $"<Relationship Id=\"{r.Item1}\" Type=\"{r.Item2}\" Target=\"{r.Item3}\"/>")) +
                "</Relationships>");
        }
        return ms.ToArray();
    }

    public MemoryStream BuildStream() => new(Build());

    private string AddStory(string type, string root, string innerXml)
    {
        var id = NextId();
        var file = $"{type}{_nextId}.xml";
        _parts[$"word/{file}"] = Encoding.UTF8.GetBytes($"<w:{root} {Namespaces}>{innerXml}</w:{root}>");
        _relationships.Add((id, $"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{type}", file));
        return id;
    }

    private string NextId() => $"rId{100 + _nextId++}";

    private static void Write(ZipArchive zip, string name, string content)
    {
        using var s = zip.CreateEntry(name).Open();
        s.Write(Encoding.UTF8.GetBytes(content));
    }

    // ------------------------------------------------------------ small markup helpers

    public static string P(string runs, string pPr = "") => $"<w:p>{(pPr.Length > 0 ? $"<w:pPr>{pPr}</w:pPr>" : "")}{runs}</w:p>";

    public static string R(string text, string rPr = "") =>
        $"<w:r>{(rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : "")}<w:t xml:space=\"preserve\">{System.Security.SecurityElement.Escape(text)}</w:t></w:r>";

    public static string Section(int widthTwips = 12240, int heightTwips = 15840, string extra = "", int margin = 1440) =>
        $"<w:sectPr>{extra}<w:pgSz w:w=\"{widthTwips}\" w:h=\"{heightTwips}\"{(widthTwips > heightTwips ? " w:orient=\"landscape\"" : "")}/>" +
        $"<w:pgMar w:top=\"{margin}\" w:right=\"{margin}\" w:bottom=\"{margin}\" w:left=\"{margin}\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>";
}