using System.IO.Compression;
using System.Security;
using System.Text;

namespace DotOffice.Tests.Xlsx;

/// <summary>
/// Test-only XLSX writer: callers provide raw SpreadsheetML fragments, the builder adds the package plumbing
/// (content types, relationships, workbook part, namespace declarations).
/// </summary>
internal sealed class XlsxBuilder
{
    public const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    private sealed record SheetSpec(string Name, string Head, string Rows, string Tail, string State);

    private readonly List<SheetSpec> _sheets = [];
    private readonly Dictionary<string, byte[]> _parts = [];

    /// <summary>Inner XML of sharedStrings.xml (si elements).</summary>
    public string? SharedStrings { get; set; }

    /// <summary>Inner XML of styles.xml (numFmts, fonts, fills, borders, cellXfs…).</summary>
    public string? Styles { get; set; }

    /// <summary>Inner XML of the a:clrScheme of theme1.xml.</summary>
    public string? ThemeColors { get; set; }

    /// <summary>Inner XML of the workbook's definedNames element (definedName elements).</summary>
    public string? DefinedNames { get; set; }

    public bool Date1904 { get; set; }
    public string? Title { get; set; }

    /// <summary>Adds a worksheet. <paramref name="rows"/> is the inner XML of sheetData.</summary>
    public XlsxBuilder AddSheet(string name, string rows = "", string head = "", string tail = "", string state = "visible")
    {
        _sheets.Add(new SheetSpec(name, head, rows, tail, state));
        return this;
    }

    /// <summary>Adds a chart sheet entry (a sheet whose part is a chartsheet, not a worksheet).</summary>
    public XlsxBuilder AddChartsheet(string name)
    {
        _sheets.Add(new SheetSpec(name, "", "", "", "chartsheet"));
        return this;
    }

    /// <summary>Adds an arbitrary package part (path without a leading slash).</summary>
    public XlsxBuilder AddPart(string path, string xml)
    {
        _parts[path] = Encoding.UTF8.GetBytes(xml);
        return this;
    }

    /// <summary>Adds a binary package part such as an image (path without a leading slash).</summary>
    public XlsxBuilder AddBinaryPart(string path, byte[] data)
    {
        _parts[path] = data;
        return this;
    }

    /// <summary>The element a worksheet needs (in its tail) to use the drawing added by <see cref="AddDrawing"/>.</summary>
    public const string DrawingElement = "<drawing r:id=\"rId1\"/>";

    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string DmlNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Adds drawing part number <paramref name="sheetNumber"/> (one-based, matching the sheet) holding <paramref name="anchors"/>,
    /// with the given media files referred to as rId1, rId2… in order. The sheet must contain <see cref="DrawingElement"/>.
    /// </summary>
    public XlsxBuilder AddDrawing(int sheetNumber, string anchors, params byte[][] media)
    {
        AddPart($"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
        AddPart($"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{RelNs}\">{anchors}</xdr:wsDr>");
        var rels = new StringBuilder();
        for (int i = 0; i < media.Length; i++)
        {
            AddBinaryPart($"xl/media/image{sheetNumber}_{i + 1}.png", media[i]);
            rels.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"{RelType}image\" Target=\"../media/image{sheetNumber}_{i + 1}.png\"/>");
        }
        AddPart($"xl/drawings/_rels/drawing{sheetNumber}.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>");
        return this;
    }

    /// <summary>The element a worksheet needs (in its tail) to use the pictures added by <see cref="AddHeaderFooterPictures"/>.</summary>
    public const string HeaderFooterPicturesElement = "<legacyDrawingHF r:id=\"rId2\"/>";

    /// <summary>
    /// Adds a legacy VML drawing for the header/footer pictures (<c>&amp;G</c>) of sheet <paramref name="sheetNumber"/>: one shape per
    /// (section id such as <c>LH</c>, style such as <c>width:54pt;height:30pt</c>), all showing <paramref name="image"/>.
    /// The sheet must contain <see cref="HeaderFooterPicturesElement"/>.
    /// </summary>
    public XlsxBuilder AddHeaderFooterPictures(int sheetNumber, byte[] image, params (string Id, string Style)[] shapes)
    {
        AddPart($"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId2\" Type=\"{RelType}vmlDrawing\" Target=\"../drawings/vmlDrawing{sheetNumber}.vml\"/></Relationships>");
        AddPart($"xl/drawings/vmlDrawing{sheetNumber}.vml",
            "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">" +
            string.Concat(shapes.Select(x => $"<v:shape id=\"{x.Id}\" style=\"position:absolute;{x.Style}\"><v:imagedata o:relid=\"rId1\" o:title=\"logo\"/></v:shape>")) + "</xml>");
        AddPart($"xl/drawings/_rels/vmlDrawing{sheetNumber}.vml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}image\" Target=\"../media/hf{sheetNumber}.png\"/></Relationships>");
        AddBinaryPart($"xl/media/hf{sheetNumber}.png", image);
        return this;
    }

    /// <summary>Like <see cref="AddDrawing"/>, with chart parts (<c>c:chartSpace</c> documents) referred to as rId1, rId2… in order.</summary>
    public XlsxBuilder AddChartDrawing(int sheetNumber, string anchors, params string[] charts)
    {
        AddPart($"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{RelType}drawing\" Target=\"../drawings/drawing{sheetNumber}.xml\"/></Relationships>");
        AddPart($"xl/drawings/drawing{sheetNumber}.xml", $"<xdr:wsDr xmlns:xdr=\"{DrawingNs}\" xmlns:a=\"{DmlNs}\" xmlns:r=\"{RelNs}\">{anchors}</xdr:wsDr>");
        var rels = new StringBuilder();
        for (int i = 0; i < charts.Length; i++)
        {
            AddPart($"xl/charts/chart{sheetNumber}_{i + 1}.xml", charts[i]);
            rels.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"{RelType}chart\" Target=\"../charts/chart{sheetNumber}_{i + 1}.xml\"/>");
        }
        AddPart($"xl/drawings/_rels/drawing{sheetNumber}.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>");
        return this;
    }

    /// <summary>A chart frame referring to relationship <paramref name="embed"/> of the drawing.</summary>
    public static string ChartRef(string embed = "rId1") =>
        "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"4\" name=\"Chart\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>" +
        $"<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"{embed}\"/></a:graphicData></a:graphic></xdr:graphicFrame>";

    private const long Emu = 12700;

    private static string Point(string tag, int column, int row, double columnOffsetPoints = 0, double rowOffsetPoints = 0) =>
        $"<xdr:{tag}><xdr:col>{column}</xdr:col><xdr:colOff>{(long)(columnOffsetPoints * Emu)}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{(long)(rowOffsetPoints * Emu)}</xdr:rowOff></xdr:{tag}>";

    /// <summary>A picture element referring to relationship <paramref name="embed"/>.</summary>
    public static string Pic(string embed = "rId1", string? srcRect = null, int rotation = 0, string flip = "", string nameAttributes = "") =>
        $"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"Picture\"{nameAttributes}/><xdr:cNvPicPr/></xdr:nvPicPr>" +
        $"<xdr:blipFill><a:blip r:embed=\"{embed}\"/>{(srcRect is null ? "" : $"<a:srcRect {srcRect}/>")}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>" +
        $"<xdr:spPr><a:xfrm rot=\"{rotation}\"{flip}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>";

    public static string TwoCellAnchor(int fromColumn, int fromRow, int toColumn, int toRow, string content, double fromColumnOffset = 0, double fromRowOffset = 0,
        double toColumnOffset = 0, double toRowOffset = 0, string clientData = "<xdr:clientData/>") =>
        $"<xdr:twoCellAnchor>{Point("from", fromColumn, fromRow, fromColumnOffset, fromRowOffset)}{Point("to", toColumn, toRow, toColumnOffset, toRowOffset)}{content}{clientData}</xdr:twoCellAnchor>";

    public static string OneCellAnchor(int column, int row, double widthPoints, double heightPoints, string content, double columnOffset = 0, double rowOffset = 0) =>
        $"<xdr:oneCellAnchor>{Point("from", column, row, columnOffset, rowOffset)}<xdr:ext cx=\"{(long)(widthPoints * Emu)}\" cy=\"{(long)(heightPoints * Emu)}\"/>{content}<xdr:clientData/></xdr:oneCellAnchor>";

    public static string AbsoluteAnchor(double xPoints, double yPoints, double widthPoints, double heightPoints, string content) =>
        $"<xdr:absoluteAnchor><xdr:pos x=\"{(long)(xPoints * Emu)}\" y=\"{(long)(yPoints * Emu)}\"/><xdr:ext cx=\"{(long)(widthPoints * Emu)}\" cy=\"{(long)(heightPoints * Emu)}\"/>{content}<xdr:clientData/></xdr:absoluteAnchor>";

    /// <summary>A shape with a text body, a connector-free chart frame etc. are all "not a picture" to the reader.</summary>
    public const string Shape = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Rectangle\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp>";

    /// <summary>A shape the reader cannot draw (custom geometry).</summary>
    public const string CustomShape = "<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Freeform\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:custGeom/></xdr:spPr></xdr:sp>";

    /// <summary>A shape with the given <c>spPr</c> content, optional <c>style</c> and <c>txBody</c> elements.</summary>
    public static string Sp(string spPr = "", string style = "", string txBody = "", string nvAttributes = "") =>
        $"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Shape\"{nvAttributes}/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>{spPr}</xdr:spPr>{style}{txBody}</xdr:sp>";

    public const string ChartFrame = "<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"4\" name=\"Chart\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>" +
                                     "<xdr:xfrm/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/></a:graphic></xdr:graphicFrame>";

    public byte[] Build()
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var rels = new StringBuilder();
            var overrides = new StringBuilder();
            int rid = 1;
            var sheetsXml = new StringBuilder();
            for (int i = 0; i < _sheets.Count; i++)
            {
                var sheet = _sheets[i];
                int id = rid++;
                bool chart = sheet.State == "chartsheet";
                Write(zip, $"xl/worksheets/sheet{i + 1}.xml", chart
                    ? $"<chartsheet xmlns=\"{MainNs}\"/>"
                    : $"<worksheet xmlns=\"{MainNs}\" xmlns:r=\"{RelNs}\">{sheet.Head}<sheetData>{sheet.Rows}</sheetData>{sheet.Tail}</worksheet>");
                rels.Append($"<Relationship Id=\"rId{id}\" Type=\"{RelType}{(chart ? "chartsheet" : "worksheet")}\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
                var state = sheet.State is "visible" or "chartsheet" ? "" : $" state=\"{sheet.State}\"";
                sheetsXml.Append($"<sheet name=\"{SecurityElement.Escape(sheet.Name)}\" sheetId=\"{i + 1}\"{state} r:id=\"rId{id}\"/>");
            }
            void Optional(string? xml, string root, string file, string type, string? extraAttributes = null)
            {
                if (xml is null) return;
                Write(zip, $"xl/{file}", $"<{root} xmlns=\"{MainNs}\"{extraAttributes}>{xml}</{root}>");
                rels.Append($"<Relationship Id=\"rId{rid++}\" Type=\"{RelType}{type}\" Target=\"{file}\"/>");
            }
            Optional(SharedStrings, "sst", "sharedStrings.xml", "sharedStrings");
            Optional(Styles, "styleSheet", "styles.xml", "styles");
            if (ThemeColors is not null)
            {
                Write(zip, "xl/theme/theme1.xml",
                    "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"><a:themeElements>" +
                    $"<a:clrScheme name=\"T\">{ThemeColors}</a:clrScheme></a:themeElements></a:theme>");
                rels.Append($"<Relationship Id=\"rId{rid++}\" Type=\"{RelType}theme\" Target=\"theme/theme1.xml\"/>");
            }
            foreach (var (name, data) in _parts)
            {
                using var s = zip.CreateEntry(name).Open();
                s.Write(data);
            }

            Write(zip, "xl/workbook.xml",
                $"<workbook xmlns=\"{MainNs}\" xmlns:r=\"{RelNs}\">" +
                (Date1904 ? "<workbookPr date1904=\"1\"/>" : "<workbookPr/>") +
                $"<sheets>{sheetsXml}</sheets>" +
                (DefinedNames is null ? "" : $"<definedNames>{DefinedNames}</definedNames>") +
                "</workbook>");
            Write(zip, "xl/_rels/workbook.xml.rels",
                $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>");
            Write(zip, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                $"{overrides}</Types>");
            var packageRels = $"<Relationship Id=\"rId1\" Type=\"{RelType}officeDocument\" Target=\"xl/workbook.xml\"/>";
            if (Title is not null)
            {
                packageRels += "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
                Write(zip, "docProps/core.xml",
                    "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                    $"xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>{SecurityElement.Escape(Title)}</dc:title><dc:creator>Tests</dc:creator></cp:coreProperties>");
            }
            Write(zip, "_rels/.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{packageRels}</Relationships>");
        }
        return ms.ToArray();
    }

    public MemoryStream BuildStream() => new(Build());

    private static void Write(ZipArchive zip, string name, string content)
    {
        using var s = zip.CreateEntry(name).Open();
        s.Write(Encoding.UTF8.GetBytes(content));
    }

    // ------------------------------------------------------------ conditional formatting helpers

    public const string X14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    /// <summary>The styles.xml parts (everything but cellXfs) conditional-formatting tests share: one default font, fill and border, and the given dxfs.</summary>
    public static string StylesWithDxfs(params string[] dxfs) =>
        "<fonts><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills><fill><patternFill patternType=\"none\"/></fill></fills>" +
        "<borders><border/></borders><cellXfs><xf fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>" +
        $"<dxfs count=\"{dxfs.Length}\">{string.Concat(dxfs)}</dxfs>";

    /// <summary>A dxf with a solid fill (bgColor, as Excel writes it) and optionally font properties.</summary>
    public static string Dxf(string? fill = null, string font = "", string border = "", string numFmt = "") =>
        $"<dxf>{(font.Length == 0 ? "" : $"<font>{font}</font>")}{numFmt}" +
        (fill is null ? "" : $"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{fill}\"/><bgColor rgb=\"FF{fill}\"/></patternFill></fill>") +
        $"{(border.Length == 0 ? "" : $"<border>{border}</border>")}</dxf>";

    /// <summary>A <c>conditionalFormatting</c> element (for the sheet's tail).</summary>
    public static string ConditionalFormatting(string sqref, params string[] rules) =>
        $"<conditionalFormatting sqref=\"{sqref}\">{string.Concat(rules)}</conditionalFormatting>";

    public static string CfRule(string type, int priority, string attributes = "", string inner = "", int? dxfId = null, bool stopIfTrue = false) =>
        $"<cfRule type=\"{type}\" priority=\"{priority}\"{(dxfId is null ? "" : $" dxfId=\"{dxfId}\"")}{(stopIfTrue ? " stopIfTrue=\"1\"" : "")} {attributes}>{inner}</cfRule>";

    public static string CellIs(string op, string operand, int priority, int dxfId, string? second = null, bool stopIfTrue = false) =>
        CfRule("cellIs", priority, $"operator=\"{op}\"", $"<formula>{operand}</formula>{(second is null ? "" : $"<formula>{second}</formula>")}", dxfId, stopIfTrue);

    public static string Expression(string formula, int priority, int dxfId, bool stopIfTrue = false) =>
        CfRule("expression", priority, "", $"<formula>{System.Security.SecurityElement.Escape(formula)}</formula>", dxfId, stopIfTrue);

    public static string Cfvo(string type, string? value = null) => $"<cfvo type=\"{type}\"{(value is null ? "" : $" val=\"{value}\"")}/>";

    public static string Rgb(string hex) => $"<color rgb=\"FF{hex}\"/>";

    /// <summary>A data bar rule; <paramref name="extensionId"/> links it to the extension written by <see cref="DataBarExtension"/>.</summary>
    public static string DataBar(int priority, string color = "638EC6", string min = "<cfvo type=\"min\"/>", string max = "<cfvo type=\"max\"/>",
        string attributes = "", string? extensionId = null) =>
        CfRule("dataBar", priority, "", $"<dataBar {attributes}>{min}{max}{Rgb(color)}</dataBar>" +
            (extensionId is null ? "" : $"<extLst><ext uri=\"{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}\" xmlns:x14=\"{X14Ns}\"><x14:id>{extensionId}</x14:id></ext></extLst>"));

    /// <summary>The worksheet <c>extLst</c> holding the Excel 2010 half of data bar rules (id, gradient flag, other attributes).</summary>
    public static string DataBarExtension(string id, bool gradient, string attributes = "", string children = "") =>
        $"<extLst><ext uri=\"{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}\" xmlns:x14=\"{X14Ns}\"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
        $"<x14:cfRule type=\"dataBar\" id=\"{id}\"><x14:dataBar minLength=\"0\" maxLength=\"100\" gradient=\"{(gradient ? 1 : 0)}\" {attributes}>" +
        $"<x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/>{children}</x14:dataBar></x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";

    public static string IconSet(int priority, string name, string thresholds, string attributes = "", string? extensionId = null) =>
        CfRule("iconSet", priority, "", $"<iconSet iconSet=\"{name}\" {attributes}>{thresholds}</iconSet>" +
            (extensionId is null ? "" : $"<extLst><ext uri=\"{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}\" xmlns:x14=\"{X14Ns}\"><x14:id>{extensionId}</x14:id></ext></extLst>"));

    /// <summary>The Excel 2010 half of an icon-set rule with custom icons: one <c>cfIcon</c> per level as (set, icon id).</summary>
    public static string IconSetExtension(string id, string name, params (string Set, int Id)[] icons) =>
        $"<extLst><ext uri=\"{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}\" xmlns:x14=\"{X14Ns}\"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
        $"<x14:cfRule type=\"iconSet\" priority=\"1\" id=\"{id}\"><x14:iconSet iconSet=\"{name}\" custom=\"1\">{string.Concat(Enumerable.Range(0, icons.Length).Select(i => "<x14:cfvo type=\"percent\"><xm:f>" + (i * 100 / icons.Length) + "</xm:f></x14:cfvo>"))}" +
        $"{string.Concat(icons.Select(c => $"<x14:cfIcon iconSet=\"{c.Set}\" iconId=\"{c.Id}\"/>"))}</x14:iconSet></x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";

    // ------------------------------------------------------------ small markup helpers

    public static string Row(int row, params string[] cells) => $"<row r=\"{row}\">{string.Concat(cells)}</row>";

    /// <summary>A numeric cell (or any cell with the given type and raw value).</summary>
    public static string C(string reference, string value, string? type = null, int? style = null) =>
        $"<c r=\"{reference}\"{(type is null ? "" : $" t=\"{type}\"")}{(style is null ? "" : $" s=\"{style}\"")}><v>{SecurityElement.Escape(value)}</v></c>";

    /// <summary>A shared-string cell referring to the given index.</summary>
    public static string S(string reference, int index, int? style = null) => C(reference, index.ToString(), "s", style);

    /// <summary>An inline-string cell.</summary>
    public static string Inline(string reference, string text, int? style = null) =>
        $"<c r=\"{reference}\" t=\"inlineStr\"{(style is null ? "" : $" s=\"{style}\"")}><is><t xml:space=\"preserve\">{SecurityElement.Escape(text)}</t></is></c>";
}
