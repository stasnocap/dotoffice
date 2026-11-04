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
