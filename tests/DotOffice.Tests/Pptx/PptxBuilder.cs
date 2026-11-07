using System.IO.Compression;
using System.Text;

namespace DotOffice.Tests.Pptx;

/// <summary>
/// Test-only PPTX writer: callers provide raw PresentationML fragments (shape trees, text styles, backgrounds), the builder adds the package
/// plumbing (content types, relationships, presentation part, one master and one layout, namespace declarations).
/// </summary>
internal sealed class PptxBuilder
{
    private const string Ns =
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"";
    private const string RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    private sealed record SlideSpec(string Tree, string Attributes, string Head);

    private readonly List<SlideSpec> _slides = [];
    private readonly Dictionary<string, byte[]> _media = [];

    /// <summary>Slide size in EMU (default 10 × 7.5 in).</summary>
    public (long Cx, long Cy) Size { get; set; } = (9_144_000, 6_858_000);

    /// <summary>Inner XML of the master's spTree after the group properties (its shapes and placeholders).</summary>
    public string MasterTree { get; set; } = "";

    /// <summary>Inner XML of the master's cSld before the spTree (a <c>p:bg</c>).</summary>
    public string MasterBackground { get; set; } = "";

    /// <summary>Attributes of the master's clrMap.</summary>
    public string ColorMap { get; set; } =
        "bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"";

    /// <summary>Inner XML of the master's txStyles (titleStyle, bodyStyle, otherStyle).</summary>
    public string TextStyles { get; set; } = "";

    /// <summary>Inner XML of the layout's spTree after the group properties.</summary>
    public string LayoutTree { get; set; } = "";

    /// <summary>Attributes of the layout's root element (such as <c>showMasterSp="0"</c>).</summary>
    public string LayoutAttributes { get; set; } = "";

    /// <summary>Inner XML of the presentation's defaultTextStyle.</summary>
    public string DefaultTextStyle { get; set; } = "";

    /// <summary>The theme's major and minor Latin fonts.</summary>
    /// <summary>The <c>a:tblStyle</c> elements of <c>ppt/tableStyles.xml</c>; empty writes no table styles part.</summary>
    public string TableStyles { get; set; } = "";

    public (string Major, string Minor) Fonts { get; set; } = ("Major Face", "Minor Face");

    /// <summary>Adds a slide. <paramref name="tree"/> is the inner XML of its spTree after the group properties; <paramref name="head"/> goes before it in cSld.</summary>
    public PptxBuilder AddSlide(string tree, string attributes = "", string head = "")
    {
        _slides.Add(new SlideSpec(tree, attributes, head));
        return this;
    }

    /// <summary>Adds a picture that every slide can refer to as relationship <c>rIdImg</c>.</summary>
    public PptxBuilder AddImage(byte[] png)
    {
        _media["ppt/media/image1.png"] = png;
        return this;
    }

    public static string Tree(string content) =>
        "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>" + content;

    /// <summary>A shape; <paramref name="placeholder"/> is the inner attributes of <c>p:ph</c>, or null for an ordinary shape.</summary>
    public static string Sp(string spPr, string? txBody = null, string? placeholder = null, string style = "") =>
        "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"s\"/><p:cNvSpPr/><p:nvPr>" + (placeholder is null ? "" : $"<p:ph {placeholder}/>") + "</p:nvPr></p:nvSpPr>" +
        $"<p:spPr>{spPr}</p:spPr>{style}" + (txBody is null ? "" : $"<p:txBody><a:bodyPr/><a:lstStyle/>{txBody}</p:txBody>") + "</p:sp>";

    /// <summary>A table frame at the slide's corner with these grid columns, table properties and rows (<c>a:tr</c> elements).</summary>
    public static string Table(long[] columns, string tblPr, string rows) =>
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"6\" name=\"t\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>" +
        Xfrm(0, 0, columns.Sum(), 1270000).Replace("a:xfrm", "p:xfrm") +
        $"<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr {tblPr}/><a:tblGrid>" +
        string.Concat(columns.Select(w => $"<a:gridCol w=\"{w}\"/>")) + $"</a:tblGrid>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>";

    /// <summary>A table cell with one paragraph of text.</summary>
    public static string Tc(string text, string attributes = "", string tcPrAttributes = "", string tcPr = "") =>
        $"<a:tc {attributes}><a:txBody><a:bodyPr/><a:lstStyle/>{P(text)}</a:txBody><a:tcPr {tcPrAttributes}>{tcPr}</a:tcPr></a:tc>";

    public static string Xfrm(long x, long y, long cx, long cy) => $"<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>";

    public static string P(string text, string pPr = "", string rPr = "") => $"<a:p>{pPr}<a:r><a:rPr lang=\"en-US\" {rPr}/><a:t>{text}</a:t></a:r></a:p>";

    public MemoryStream BuildStream()
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var types = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Default Extension=\"png\" ContentType=\"image/png\"/></Types>");
            Write(zip, "[Content_Types].xml", types.ToString());
            Write(zip, "_rels/.rels", Rels(("rId1", "officeDocument", "ppt/presentation.xml")));

            var ids = new StringBuilder();
            var presentationRels = new List<(string, string, string)> { ("rIdM", "slideMaster", "slideMasters/slideMaster1.xml") };
            for (int i = 0; i < _slides.Count; i++)
            {
                ids.Append($"<p:sldId id=\"{256 + i}\" r:id=\"rIdS{i + 1}\"/>");
                presentationRels.Add(($"rIdS{i + 1}", "slide", $"slides/slide{i + 1}.xml"));
            }
            Write(zip, "ppt/presentation.xml",
                $"<p:presentation {Ns}><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rIdM\"/></p:sldMasterIdLst><p:sldIdLst>{ids}</p:sldIdLst>" +
                $"<p:sldSz cx=\"{Size.Cx}\" cy=\"{Size.Cy}\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/><p:defaultTextStyle>{DefaultTextStyle}</p:defaultTextStyle></p:presentation>");
            if (TableStyles.Length > 0)
            {
                presentationRels.Add(("rIdTS", "tableStyles", "tableStyles.xml"));
                Write(zip, "ppt/tableStyles.xml", $"<a:tblStyleLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" def=\"\">{TableStyles}</a:tblStyleLst>");
            }
            Write(zip, "ppt/_rels/presentation.xml.rels", Rels([.. presentationRels]));

            Write(zip, "ppt/slideMasters/slideMaster1.xml",
                $"<p:sldMaster {Ns}><p:cSld>{MasterBackground}<p:spTree>{Tree(MasterTree)}</p:spTree></p:cSld><p:clrMap {ColorMap}/>" +
                $"<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rIdL\"/></p:sldLayoutIdLst><p:txStyles>{TextStyles}</p:txStyles></p:sldMaster>");
            Write(zip, "ppt/slideMasters/_rels/slideMaster1.xml.rels", Rels(("rIdL", "slideLayout", "../slideLayouts/slideLayout1.xml"), ("rIdT", "theme", "../theme/theme1.xml"), ("rIdImg", "image", "../media/image1.png")));
            Write(zip, "ppt/slideLayouts/slideLayout1.xml", $"<p:sldLayout {Ns} {LayoutAttributes}><p:cSld><p:spTree>{Tree(LayoutTree)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>");
            Write(zip, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", Rels(("rIdM", "slideMaster", "../slideMasters/slideMaster1.xml")));
            Write(zip, "ppt/theme/theme1.xml",
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"><a:themeElements><a:clrScheme name=\"C\">" +
                "<a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>" +
                "<a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3>" +
                "<a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6>" +
                "<a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme>" +
                $"<a:fontScheme name=\"F\"><a:majorFont><a:latin typeface=\"{Fonts.Major}\"/></a:majorFont><a:minorFont><a:latin typeface=\"{Fonts.Minor}\"/></a:minorFont></a:fontScheme>" +
                "</a:themeElements></a:theme>");

            for (int i = 0; i < _slides.Count; i++)
            {
                var slide = _slides[i];
                Write(zip, $"ppt/slides/slide{i + 1}.xml", $"<p:sld {Ns} {slide.Attributes}><p:cSld>{slide.Head}<p:spTree>{Tree(slide.Tree)}</p:spTree></p:cSld></p:sld>");
                Write(zip, $"ppt/slides/_rels/slide{i + 1}.xml.rels", Rels(("rIdL", "slideLayout", "../slideLayouts/slideLayout1.xml"), ("rIdImg", "image", "../media/image1.png")));
            }
            foreach (var (name, bytes) in _media)
            {
                using var s = zip.CreateEntry(name).Open();
                s.Write(bytes);
            }
        }
        ms.Position = 0;
        return ms;
    }

    private static string Rels(params (string Id, string Type, string Target)[] rels) =>
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        string.Concat(rels.Select(r => $"<Relationship Id=\"{r.Id}\" Type=\"{RelType}{r.Type}\" Target=\"{r.Target}\"/>")) + "</Relationships>";

    private static void Write(ZipArchive zip, string name, string content)
    {
        using var s = zip.CreateEntry(name).Open();
        s.Write(Encoding.UTF8.GetBytes(content));
    }
}
