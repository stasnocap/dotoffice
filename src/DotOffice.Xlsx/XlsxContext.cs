using System.Globalization;
using System.Text;
using System.Xml;
using DotOffice.Core;
using DotOffice.Opc;

namespace DotOffice.Xlsx;

/// <summary>A sheet entry of workbook.xml.</summary>
internal sealed record SheetEntry(string Name, string? RelationshipId, SheetVisibility Visibility);

/// <summary>A defined name of workbook.xml, optionally scoped to one sheet (by position).</summary>
internal sealed record DefinedName(string Name, int? LocalSheetId, string Formula);

/// <summary>Package-wide state while reading one workbook.</summary>
internal sealed class XlsxContext(OpcPackage package, XlsxReaderOptions options, string workbookPart)
{
    private readonly List<RichString> _sharedStrings = [];
    private List<CellFormat> _formats = [];
    private Func<System.Xml.Linq.XElement?, Color?> _resolveColor = _ => null;
    private List<Dxf> _dxfs = [];
    private IReadOnlyList<Color?>? _palette;
    private Func<System.Xml.Linq.XElement, Dxf> _readInlineDxf = _ => new Dxf();
    private CfSheet _conditional = new();
    private long _cellCount;

    public OpcPackage Package { get; } = package;

    public XlsxReaderOptions Options { get; } = options;

    public string WorkbookPart { get; } = workbookPart;

    /// <summary>True when the workbook uses the 1904 date system.</summary>
    public bool Date1904 { get; private set; }

    public Workbook ReadWorkbook()
    {
        var workbook = new Workbook();
        ReadCoreProperties(workbook);

        var sheets = new List<SheetEntry>();
        var names = new List<DefinedName>();
        using (var r = Package.OpenXml(WorkbookPart))
            ReadWorkbookXml(r, sheets, names);

        ReadStyles();

        var shared = Package.FindRelationship(WorkbookPart, "sharedStrings");
        if (shared is not null && Package.Exists(shared.Target))
        {
            using var r = Package.OpenXml(shared.Target);
            ReadSharedStrings(r);
        }

        var rels = Package.GetRelationships(WorkbookPart);
        for (int index = 0; index < sheets.Count; index++)
        {
            var entry = sheets[index];
            if (entry.RelationshipId is null || !rels.TryGetValue(entry.RelationshipId, out var rel) || rel.IsExternal)
            {
                Options.Warnings.Warn("xlsx.missing-sheet", $"Sheet '{entry.Name}' has no worksheet part and was skipped.");
                continue;
            }
            if (rel.Kind != "worksheet")
            {
                Options.Warnings.Warn("xlsx.unsupported-sheet", $"Sheet '{entry.Name}' is a {rel.Kind}, which is not supported, and was skipped.");
                continue;
            }
            if (!Package.Exists(rel.Target))
            {
                Options.Warnings.Warn("xlsx.missing-sheet", $"Worksheet part '{rel.Target}' of sheet '{entry.Name}' is missing.");
                continue;
            }

            var sheet = new Worksheet(entry.Name) { Visibility = entry.Visibility };
            _sheetWarnings.Clear();
            _conditional = new CfSheet();
            int position = index;
            CfParser.NameResolver = word =>
                (names.FirstOrDefault(n => n.LocalSheetId == position && string.Equals(n.Name, word, StringComparison.OrdinalIgnoreCase))
                 ?? names.FirstOrDefault(n => n.LocalSheetId is null && string.Equals(n.Name, word, StringComparison.OrdinalIgnoreCase)))?.Formula;
            try
            {
                using var r = Package.OpenXml(rel.Target);
                ReadWorksheet(r, sheet);
            }
            finally { CfParser.NameResolver = null; }
            ReadDrawing(sheet, rel.Target);
            ReadComments(sheet, rel.Target);
            WarnAboutDeferredFeatures(sheet, rel.Target);
            ApplyDefinedNames(sheet, index, names);
            _conditionalJobs.Add((sheet, _conditional));
            workbook.Sheets.Add(sheet);
        }
        // Conditional-formatting formulas may read other sheets, so the rules are applied once every sheet has been read.
        foreach (var (sheet, rules) in _conditionalJobs)
        {
            _sheetWarnings.Clear();
            ConditionalFormatting.Apply(sheet, rules,
                new ConditionalFormatting.Environment(_dxfs, Date1904, _palette, (Options.Today ?? DateTime.Today).Date,
                    name => workbook.Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))),
                message => WarnOnce(sheet, "xlsx.conditional-formatting", message));
        }
        // Charts may refer to cells of any sheet, so they are built once every sheet has been read.
        foreach (var (sheet, build) in _chartJobs)
        {
            try { sheet.Charts.Add(build(workbook.Sheets)); }
            catch (Exception e) when (e is FormatException or InvalidOperationException or ArgumentException or XmlException or OverflowException)
            {
                Options.Warnings.Warn("xlsx.chart", $"Sheet '{sheet.Name}' has a chart that could not be read ({e.Message}); it is not rendered.");
            }
        }
        return workbook;
    }

    private readonly List<(Worksheet Sheet, Func<IReadOnlyList<Worksheet>, SheetChart> Build)> _chartJobs = [];
    private readonly List<(Worksheet Sheet, CfSheet Rules)> _conditionalJobs = [];

    internal void DeferChart(Worksheet sheet, Func<IReadOnlyList<Worksheet>, SheetChart> build) => _chartJobs.Add((sheet, build));

    // ------------------------------------------------------------------ workbook.xml

    private void ReadWorkbookXml(XmlReader r, List<SheetEntry> sheets, List<DefinedName> names)
    {
        r.MoveToContent();
        while (!r.EOF)
        {
            if (r.NodeType != XmlNodeType.Element)
            {
                r.Read();
                continue;
            }
            switch (r.LocalName)
            {
                case "workbookPr":
                    Date1904 = r.BoolAttr("date1904");
                    r.Read();
                    break;
                case "sheet":
                    var name = r.GetAttribute("name");
                    if (string.IsNullOrEmpty(name)) { r.Read(); break; }
                    var visibility = r.GetAttribute("state") switch
                    {
                        "hidden" => SheetVisibility.Hidden,
                        "veryHidden" => SheetVisibility.VeryHidden,
                        _ => SheetVisibility.Visible,
                    };
                    sheets.Add(new SheetEntry(name, r.RelationshipId(), visibility));
                    r.Read();
                    break;
                case "definedName":
                    var definedName = r.GetAttribute("name");
                    var local = r.IntAttr("localSheetId");
                    if (definedName is null || r.IsEmptyElement) { r.Read(); break; }
                    names.Add(new DefinedName(definedName, local, r.ReadElementContentAsString()));
                    break;
                default:
                    r.Read();
                    break;
            }
        }
    }

    private void ReadCoreProperties(Workbook workbook)
    {
        var rel = Package.FindRelationship("", "core-properties");
        if (rel is null || !Package.Exists(rel.Target)) return;
        try
        {
            using var r = Package.OpenXml(rel.Target);
            r.MoveToContent();
            while (!r.EOF)
            {
                if (r.NodeType == XmlNodeType.Element && !r.IsEmptyElement && r.LocalName == "title")
                    workbook.Title = NullIfEmpty(r.ReadElementContentAsString());
                else if (r.NodeType == XmlNodeType.Element && !r.IsEmptyElement && r.LocalName == "creator")
                    workbook.Author = NullIfEmpty(r.ReadElementContentAsString());
                else
                    r.Read();
            }
        }
        catch (XmlException)
        {
            Options.Warnings.Warn("xlsx.core-properties", "The document properties part is malformed and was ignored.");
        }
    }

    private static string? NullIfEmpty(string s) => s.Length == 0 ? null : s;

    // ------------------------------------------------------------------ styles.xml and theme

    /// <summary>The workbook theme (the Office default when the file has none).</summary>
    internal XlsxTheme Theme { get; private set; } = XlsxTheme.Default;

    private void ReadStyles()
    {
        var theme = XlsxTheme.Default;
        var themeRel = Package.FindRelationship(WorkbookPart, "theme");
        if (themeRel is not null && Package.Exists(themeRel.Target))
        {
            try
            {
                using var r = Package.OpenXml(themeRel.Target);
                theme = XlsxTheme.Read(r);
            }
            catch (XmlException)
            {
                Options.Warnings.Warn("xlsx.theme", "The theme part is malformed; default theme colors are used.");
            }
        }

        Theme = theme;
        var stylesRel = Package.FindRelationship(WorkbookPart, "styles");
        if (stylesRel is null || !Package.Exists(stylesRel.Target)) return;
        using var styles = Package.OpenXml(stylesRel.Target);
        var table = XlsxStyles.Read(styles, theme, Options.Warnings);
        _formats = table.Formats;
        _resolveColor = table.Resolve;
        _dxfs = table.Dxfs;
        _palette = table.Palette;
        _readInlineDxf = table.ReadInlineDxf;
    }

    private CellFormat FormatOf(string? styleAttribute) =>
        int.TryParse(styleAttribute, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index >= 0 && index < _formats.Count
            ? _formats[index]
            : CellFormat.Default;

    // ------------------------------------------------------------------ sharedStrings.xml

    private void ReadSharedStrings(XmlReader r)
    {
        r.MoveToContent();
        while (!r.EOF)
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "si")
            {
                _sharedStrings.Add(r.ReadRichText());
            }
            else
            {
                r.Read();
            }
        }
    }

    // ------------------------------------------------------------------ worksheet

    private readonly HashSet<string> _sheetWarnings = [];
    private string? _drawingRelationship;
    private readonly Dictionary<string, ImageData> _images = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Loads an image part once per package.</summary>
    internal ImageData GetImage(string partName)
    {
        if (_images.TryGetValue(partName, out var image)) return image;
        return _images[partName] = new ImageData(Package.ReadAllBytes(partName), partName);
    }

    /// <summary>Reads the pictures of the sheet's drawing; charts, shapes and anything else in it are reported once.</summary>
    /// <summary>Reads the legacy comments part (<c>comments</c>) of a sheet; the text of each comment is its runs joined.</summary>
    private void ReadComments(Worksheet sheet, string sheetPart)
    {
        var rel = Package.FindRelationship(sheetPart, "comments");
        if (rel is null || rel.IsExternal || !Package.Exists(rel.Target)) return;
        try
        {
            System.Xml.Linq.XDocument doc;
            using (var reader = Package.OpenXml(rel.Target)) doc = System.Xml.Linq.XDocument.Load(reader);
            var authors = doc.Descendants().Where(e => e.Name.LocalName == "author").Select(e => e.Value).ToList();
            foreach (var comment in doc.Descendants().Where(e => e.Name.LocalName == "comment"))
            {
                var reference = (string?)comment.Attribute("ref") ?? "";
                if (!CellReference.TryParse(reference, out var row, out var column)) continue;
                var text = comment.Elements().FirstOrDefault(e => e.Name.LocalName == "text");
                if (text is null) continue;
                var content = string.Concat(text.Descendants().Where(e => e.Name.LocalName == "t" && e.Parent?.Name.LocalName != "rPh").Select(e => e.Value));
                string? author = int.TryParse((string?)comment.Attribute("authorId"), out var id) && id >= 0 && id < authors.Count ? authors[id] : null;
                sheet.Comments.Add(new SheetComment(row, column, reference.Replace("$", ""), string.IsNullOrWhiteSpace(author) ? null : author, content));
            }
            sheet.Comments.Sort((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Column.CompareTo(b.Column));
        }
        catch (System.Xml.XmlException)
        {
            WarnOnce(sheet, "xlsx.comments", "has a comments part that could not be read; its comments are not printed.");
        }
    }

    private void ReadDrawing(Worksheet sheet, string sheetPart)
    {
        var id = _drawingRelationship;
        _drawingRelationship = null;
        if (id is null) return;
        if (!Package.GetRelationships(sheetPart).TryGetValue(id, out var rel) || rel.IsExternal || !Package.Exists(rel.Target))
        {
            WarnOnce(sheet, "xlsx.drawing", "has a drawing whose part is missing, so its pictures, charts and shapes are not rendered.");
            return;
        }
        try
        {
            var skipped = XlsxDrawings.Read(this, sheet, rel.Target);
            if (skipped.Any)
            {
                var parts = new List<string>();
                if (skipped.Charts > 0)
                {
                    var kinds = skipped.ChartKinds?.Distinct().ToList() ?? [];
                    parts.Add(kinds.Count > 0 ? $"{(skipped.Charts == 1 ? "a" : skipped.Charts.ToString(CultureInfo.InvariantCulture))} {string.Join("/", kinds)} chart{(skipped.Charts == 1 ? "" : "s")}"
                        : skipped.Charts == 1 ? "a chart" : $"{skipped.Charts} charts");
                }
                if (skipped.Shapes > 0) parts.Add(skipped.Shapes == 1 ? "a shape with custom geometry" : $"{skipped.Shapes} shapes with custom geometry");
                if (skipped.Other > 0) parts.Add("other drawing objects");
                WarnOnce(sheet, "xlsx.drawing", $"contains {string.Join(" and ", parts)}, which {(parts.Count == 1 && skipped.Charts + skipped.Shapes <= 1 ? "is" : "are")} not rendered (pictures, shapes and column, bar, line, area, pie, doughnut, scatter, bubble and radar charts are).");
            }
            if (skipped.Notes.Count > 0)
                WarnOnce(sheet, "xlsx.drawing-approximated", $"has drawing objects that are only approximated: {string.Join("; ", skipped.Notes)}.");
        }
        catch (XmlException)
        {
            WarnOnce(sheet, "xlsx.drawing", "has a malformed drawing, so its pictures, charts and shapes are not rendered.");
        }
    }

    private void WarnOnce(Worksheet sheet, string code, string message)
    {
        if (_sheetWarnings.Add(code)) Options.Warnings.Warn(code, $"Sheet '{sheet.Name}' {message}");
    }

    /// <summary>Features that are not rendered yet are reported once per sheet and then ignored.</summary>
    private void WarnAboutDeferredFeatures(Worksheet sheet, string sheetPart)
    {
        foreach (var rel in Package.GetRelationships(sheetPart).Values)
        {
            if (rel.IsExternal) continue;
            switch (rel.Kind)
            {
                case "pivotTable":
                    ReadPivotTable(sheet, rel.Target);
                    break;
            }
        }
    }

    /// <summary>
    /// A pivot table's cells are ordinary cells of the sheet, holding the values Excel last calculated, so they print as they are. That is
    /// faithful unless the table gets its look from a pivot style or pivot-area formats (neither is applied) or is set to refresh when the
    /// file opens (Excel then shows other values); only those cases warn, and the warning says which.
    /// </summary>
    private void ReadPivotTable(Worksheet sheet, string part)
    {
        System.Xml.Linq.XElement? root = null;
        try
        {
            using var r = Package.OpenXml(part);
            root = System.Xml.Linq.XDocument.Load(r).Root;
        }
        catch (XmlException) { }
        if (root is null || root.Name.LocalName != "pivotTableDefinition")
        {
            WarnOnce(sheet, "xlsx.pivot", "contains a pivot table that could not be read; its cells are printed as they are, without pivot formatting.");
            return;
        }
        var reasons = new List<string>();
        if ((string?)root.Elements().FirstOrDefault(e => e.Name.LocalName == "pivotTableStyleInfo")?.Attribute("name") is { Length: > 0 } style)
            reasons.Add($"its table style '{style}' is not applied");
        if (root.Elements().FirstOrDefault(e => e.Name.LocalName == "formats")?.HasElements == true)
            reasons.Add("its pivot-area formats are not applied");
        var cache = Package.FindRelationship(part, "pivotCacheDefinition");
        if (cache is not null && !cache.IsExternal && Package.Exists(cache.Target))
        {
            try
            {
                using var r = Package.OpenXml(cache.Target);
                if ((string?)System.Xml.Linq.XDocument.Load(r).Root?.Attribute("refreshOnLoad") is "1" or "true")
                    reasons.Add("it is set to refresh when the file opens, so Excel may show other values");
            }
            catch (XmlException) { }
        }
        if (reasons.Count > 0)
            WarnOnce(sheet, "xlsx.pivot", $"contains the pivot table '{(string?)root.Attribute("name")}', which is printed from its cached cells; {string.Join(" and ", reasons)}.");
    }

    private void ReadWorksheet(XmlReader r, Worksheet sheet)
    {
        bool formulaWarned = false;
        sheet.DefaultColumnWidth = 48; // Excel's 8.43-character default (64 px)
        bool fitToPage = false;
        bool breaksAreRows = true;
        r.MoveToContent();
        while (!r.EOF)
        {
            if (r.NodeType != XmlNodeType.Element) { r.Read(); continue; }
            switch (r.LocalName)
            {
                case "sheetFormatPr":
                    ReadSheetFormat(r, sheet);
                    r.Read();
                    break;
                case "cols":
                    ReadColumns(r, sheet);
                    break;
                case "mergeCells":
                    ReadMergeCells(r, sheet);
                    break;
                case "sheetData":
                    ReadSheetData(r, sheet, ref formulaWarned);
                    break;
                case "drawing":
                    _drawingRelationship = r.RelationshipId();
                    r.Read();
                    break;
                case "conditionalFormatting" when IsMainNamespace(r.NamespaceURI):
                    // The x14 conditionalFormatting elements inside extLst are not rules of the main namespace; they are read with the extension list.
                    ConditionalFormatting.ReadBlock((System.Xml.Linq.XElement)System.Xml.Linq.XNode.ReadFrom(r), _conditional, _resolveColor);
                    break;
                case "extLst" when IsMainNamespace(r.NamespaceURI):
                    ConditionalFormatting.ReadExtensions((System.Xml.Linq.XElement)System.Xml.Linq.XNode.ReadFrom(r), _conditional, _resolveColor,
                        dxf => { _dxfs.Add(_readInlineDxf(dxf)); return _dxfs.Count - 1; });
                    break;
                case "pageSetUpPr":
                    fitToPage = r.BoolAttr("fitToPage");
                    r.Read();
                    break;
                case "printOptions":
                    sheet.PageSetup.CenterHorizontally = r.BoolAttr("horizontalCentered");
                    sheet.PageSetup.CenterVertically = r.BoolAttr("verticalCentered");
                    sheet.PageSetup.PrintGridLines = r.BoolAttr("gridLines");
                    r.Read();
                    break;
                case "pageMargins":
                    ReadMargins(r, sheet.PageSetup);
                    r.Read();
                    break;
                case "pageSetup":
                    ReadPageSetup(r, sheet, fitToPage);
                    r.Read();
                    break;
                case "headerFooter":
                    ReadHeaderFooter(r, sheet.PageSetup);
                    break;
                case "rowBreaks":
                case "colBreaks":
                    breaksAreRows = r.LocalName == "rowBreaks";
                    r.Read();
                    break;
                case "brk":
                    if (r.IntAttr("id") is { } breakId && breakId >= 1)
                        (breaksAreRows ? sheet.PageSetup.RowBreaks : sheet.PageSetup.ColumnBreaks).Add(breakId - 1);
                    r.Read();
                    break;
                default:
                    r.Read();
                    break;
            }
        }
    }

    private static bool IsMainNamespace(string uri) =>
        uri is "http://schemas.openxmlformats.org/spreadsheetml/2006/main" or "http://purl.oclc.org/ooxml/spreadsheetml/main";

    private static void ReadMargins(XmlReader r, SheetPageSetup setup)
    {
        double Inches(string name, double fallback) => (r.DoubleAttr(name) is { } v && v >= 0 ? v : fallback) * 72;
        setup.Margins = new Thickness(Inches("top", 0.75), Inches("right", 0.7), Inches("bottom", 0.75), Inches("left", 0.7));
        setup.HeaderDistance = Inches("header", 0.3);
        setup.FooterDistance = Inches("footer", 0.3);
    }

    private static (double, double) In(double w, double h) => (w * 72, h * 72);
    private static (double, double) Mm(double w, double h) => (w * 72 / 25.4, h * 72 / 25.4);

    /// <summary>
    /// Portrait sizes in points for every SpreadsheetML <c>paperSize</c> id that names a fixed size (the ids of the Windows
    /// <c>DEVMODE</c> list). The "rotated" ids stand for the same sheet turned sideways: the page's own orientation decides, so they
    /// take the size of their upright twin.
    /// </summary>
    private static readonly Dictionary<int, (double Width, double Height)> PaperSizes = new()
    {
        [1] = In(8.5, 11), [2] = In(8.5, 11), [3] = In(11, 17), [4] = In(17, 11), [5] = In(8.5, 14), [6] = In(5.5, 8.5),
        [7] = In(7.25, 10.5), [8] = Mm(297, 420), [9] = Mm(210, 297), [10] = Mm(210, 297), [11] = Mm(148, 210),
        [12] = Mm(257, 364), [13] = Mm(182, 257), [14] = In(8.5, 13), [15] = Mm(215, 275), [16] = In(10, 14),
        [17] = In(11, 17), [18] = In(8.5, 11), [19] = In(3.875, 8.875), [20] = In(4.125, 9.5), [21] = In(4.5, 10.375),
        [22] = In(4.75, 11), [23] = In(5, 11.5), [24] = In(17, 22), [25] = In(22, 34), [26] = In(34, 44), [27] = Mm(110, 220),
        [28] = Mm(162, 229), [29] = Mm(324, 458), [30] = Mm(229, 324), [31] = Mm(114, 162), [32] = Mm(114, 229),
        [33] = Mm(250, 353), [34] = Mm(176, 250), [35] = Mm(125, 176), [36] = Mm(110, 230), [37] = In(3.875, 7.5),
        [38] = In(3.625, 6.5), [39] = In(14.875, 11), [40] = In(8.5, 12), [41] = In(8.5, 13), [42] = Mm(250, 353),
        [43] = Mm(100, 148), [44] = In(9, 11), [45] = In(10, 11), [46] = In(11, 15), [47] = Mm(220, 220),
        [50] = In(9.5, 12), [51] = In(9.5, 15), [52] = In(11.69, 18), [53] = In(9.27, 12.69), [54] = In(8.5, 11),
        [55] = Mm(210, 297), [56] = In(9.275, 12), [57] = Mm(227, 356), [58] = Mm(305, 487), [59] = In(8.5, 12.69),
        [60] = Mm(210, 330), [61] = Mm(148, 210), [62] = Mm(182, 257), [63] = Mm(322, 445), [64] = Mm(174, 235),
        [65] = Mm(201, 276), [66] = Mm(420, 594), [67] = Mm(297, 420), [68] = Mm(322, 445), [69] = Mm(148, 200),
        [70] = Mm(105, 148), [71] = Mm(240, 332), [72] = Mm(216, 277), [73] = Mm(120, 235), [74] = Mm(90, 205),
        [75] = In(8.5, 11), [76] = Mm(297, 420), [77] = Mm(210, 297), [78] = Mm(148, 210), [79] = Mm(257, 364),
        [80] = Mm(182, 257), [81] = Mm(100, 148), [82] = Mm(148, 200), [83] = Mm(105, 148), [84] = Mm(240, 332),
        [85] = Mm(216, 277), [86] = Mm(120, 235), [87] = Mm(90, 205), [88] = Mm(128, 182), [89] = Mm(128, 182),
        [90] = In(11, 12), [91] = Mm(105, 235), [92] = Mm(105, 235), [93] = Mm(146, 215), [94] = Mm(97, 151),
        [95] = Mm(97, 151), [96] = Mm(102, 165), [97] = Mm(102, 176), [98] = Mm(125, 176), [99] = Mm(110, 208),
        [100] = Mm(110, 220), [101] = Mm(120, 230), [102] = Mm(160, 230), [103] = Mm(120, 309), [104] = Mm(229, 324),
        [105] = Mm(324, 458), [106] = Mm(146, 215), [107] = Mm(97, 151), [108] = Mm(97, 151), [109] = Mm(102, 165),
        [110] = Mm(102, 176), [111] = Mm(125, 176), [112] = Mm(110, 208), [113] = Mm(110, 220), [114] = Mm(120, 230),
        [115] = Mm(160, 230), [116] = Mm(120, 309), [117] = Mm(229, 324), [118] = Mm(324, 458),
    };

    private void ReadPageSetup(XmlReader r, Worksheet sheet, bool fitToPage)
    {
        var setup = sheet.PageSetup;
        double width, height;
        if (PaperSizes.TryGetValue(r.IntAttr("paperSize") ?? 1, out var paper)) (width, height) = paper;
        else
        {
            Options.Warnings.Warn("xlsx.paper-size", $"Sheet '{sheet.Name}' uses an unsupported paper size; Letter is used instead.");
            (width, height) = (612, 792);
        }
        bool landscape = r.GetAttribute("orientation") == "landscape";
        setup.Orientation = landscape ? Orientation.Landscape : Orientation.Portrait;
        // PageWidth/PageHeight are the effective (oriented) page size.
        (setup.PageWidth, setup.PageHeight) = landscape ? (height, width) : (width, height);
        if (r.IntAttr("scale") is { } scale) setup.Scale = Math.Clamp(scale, 10, 400) / 100.0;
        if (fitToPage)
        {
            setup.FitToWidth = Math.Max(r.IntAttr("fitToWidth") ?? 1, 0);
            setup.FitToHeight = Math.Max(r.IntAttr("fitToHeight") ?? 1, 0);
        }
        setup.PageOrderOverThenDown = r.GetAttribute("pageOrder") == "overThenDown";
        if (r.BoolAttr("useFirstPageNumber") && r.IntAttr("firstPageNumber") is { } firstNumber)
        {
            setup.UseFirstPageNumber = true;
            setup.FirstPageNumber = firstNumber;
        }
        setup.Comments = r.GetAttribute("cellComments") switch { "atEnd" => CommentPrinting.AtEnd, "asDisplayed" => CommentPrinting.AsDisplayed, _ => CommentPrinting.None };
        if (setup.Comments == CommentPrinting.AsDisplayed)
            Options.Warnings.Warn("xlsx.comments", $"Sheet '{sheet.Name}' prints its comments as displayed on the sheet; they are listed at the end instead.");
    }

    private static void ReadHeaderFooter(XmlReader r, SheetPageSetup setup)
    {
        bool different = r.BoolAttr("differentFirst");
        bool oddEven = r.BoolAttr("differentOddEven");
        setup.HeaderFooterScaleWithDocument = r.GetAttribute("scaleWithDoc") is not ("0" or "false");
        var texts = new Dictionary<string, string>();
        if (r.IsEmptyElement) { r.Read(); return; }
        int depth = r.Depth;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName is "oddHeader" or "oddFooter" or "evenHeader" or "evenFooter" or "firstHeader" or "firstFooter")
            {
                string name = r.LocalName;
                if (r.IsEmptyElement) { texts[name] = ""; r.Read(); }
                else texts[name] = r.ReadElementContentAsString();
            }
            else r.Read();
        }
        r.Read();

        HeaderFooterParts Parts(string key) => SplitHeaderFooter(texts.GetValueOrDefault(key, ""));
        setup.Header = new HeaderFooterText
        {
            Odd = Parts("oddHeader"),
            First = different ? Parts("firstHeader") : null,
            Even = oddEven ? Parts("evenHeader") : null,
        };
        setup.Footer = new HeaderFooterText
        {
            Odd = Parts("oddFooter"),
            First = different ? Parts("firstFooter") : null,
            Even = oddEven ? Parts("evenFooter") : null,
        };
    }

    /// <summary>Splits header/footer text at <c>&amp;L</c>, <c>&amp;C</c> and <c>&amp;R</c>; other codes stay in the strings. Text before any section code is centered.</summary>
    internal static HeaderFooterParts SplitHeaderFooter(string text)
    {
        var parts = new[] { new StringBuilder(), new StringBuilder(), new StringBuilder() };
        int current = 1;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c != '&' || i + 1 >= text.Length) { parts[current].Append(c); continue; }
            char next = text[i + 1];
            switch (next)
            {
                case 'L': current = 0; i++; break;
                case 'C': current = 1; i++; break;
                case 'R': current = 2; i++; break;
                case '"':
                {
                    int end = text.IndexOf('"', i + 2);
                    if (end < 0) end = text.Length - 1;
                    parts[current].Append(text, i, end - i + 1);
                    i = end;
                    break;
                }
                default:
                    parts[current].Append(c).Append(next);
                    i++;
                    break;
            }
        }
        return new HeaderFooterParts(parts[0].ToString(), parts[1].ToString(), parts[2].ToString());
    }

    private const double MaxDigitWidth = 7;

    /// <summary>Converts a width in characters of the default font's digit to points (ECMA-376 18.3.1.13).</summary>
    internal static double WidthToPoints(double characters)
    {
        double pixels = Math.Truncate((256 * characters + Math.Truncate(128 / MaxDigitWidth)) / 256 * MaxDigitWidth);
        return pixels * 0.75;
    }

    private static double? DoubleAttr(XmlReader r, string name) =>
        double.TryParse(r.GetAttribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && double.IsFinite(v) && v >= 0 ? v : null;

    private static void ReadSheetFormat(XmlReader r, Worksheet sheet)
    {
        if (DoubleAttr(r, "defaultColWidth") is { } width) sheet.DefaultColumnWidth = WidthToPoints(width);
        else if (DoubleAttr(r, "baseColWidth") is { } baseWidth)
            sheet.DefaultColumnWidth = Math.Ceiling((Math.Truncate(baseWidth * MaxDigitWidth) + 5) / 8) * 8 * 0.75;
        if (DoubleAttr(r, "defaultRowHeight") is { } height && height > 0) sheet.DefaultRowHeight = Math.Min(height, 409.5);
    }

    private void ReadColumns(XmlReader r, Worksheet sheet)
    {
        if (r.IsEmptyElement) { r.Read(); return; }
        int depth = r.Depth;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "col")
            {
                int min = (r.IntAttr("min") ?? 1) - 1;
                int max = Math.Min((r.IntAttr("max") ?? min + 1) - 1, CellReference.MaxColumns - 1);
                double? width = DoubleAttr(r, "width");
                bool hidden = r.BoolAttr("hidden");
                for (int c = Math.Max(min, 0); c <= max; c++)
                {
                    if (width is { } w) sheet.ColumnWidths[c] = Math.Min(WidthToPoints(w), 1790);
                    if (hidden) sheet.HiddenColumns.Add(c);
                }
            }
            r.Read();
        }
        r.Read();
    }

    private void ReadMergeCells(XmlReader r, Worksheet sheet)
    {
        if (r.IsEmptyElement) { r.Read(); return; }
        int depth = r.Depth;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "mergeCell")
            {
                if (r.GetAttribute("ref") is { } reference && CellReference.TryParseRange(reference, out var range))
                {
                    if (range.FirstRow != range.LastRow || range.FirstColumn != range.LastColumn) sheet.Merges.Add(range);
                }
                else Options.Warnings.Warn("xlsx.bad-reference", $"Sheet '{sheet.Name}' has an invalid merged range.");
            }
            r.Read();
        }
        r.Read();
    }

    private void ReadSheetData(XmlReader r, Worksheet sheet, ref bool formulaWarned)
    {
        if (r.IsEmptyElement)
        {
            r.Read();
            return;
        }
        int depth = r.Depth;
        int rowIndex = -1;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType == XmlNodeType.Element && r.LocalName == "row")
            {
                rowIndex = r.IntAttr("r") is { } rowNumber ? rowNumber - 1 : rowIndex + 1;
                if (rowIndex is < 0 or >= CellReference.MaxRows)
                {
                    Options.Warnings.Warn("xlsx.bad-reference", $"Sheet '{sheet.Name}' has a row outside the supported range.");
                    r.Skip();
                    continue;
                }
                if (DoubleAttr(r, "ht") is { } rowHeight) sheet.RowHeights[rowIndex] = Math.Min(rowHeight, 409.5);
                if (r.BoolAttr("hidden")) sheet.HiddenRows.Add(rowIndex);
                ReadRow(r, sheet, rowIndex, ref formulaWarned);
            }
            else
            {
                r.Read();
            }
        }
        r.Read();
    }

    /// <summary>Reads the cells of a <c>row</c>; leaves the reader after the row's end tag.</summary>
    private void ReadRow(XmlReader r, Worksheet sheet, int rowIndex, ref bool formulaWarned)
    {
        if (r.IsEmptyElement)
        {
            r.Read();
            return;
        }
        int depth = r.Depth;
        int column = -1;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType != XmlNodeType.Element || r.LocalName != "c")
            {
                r.Read();
                continue;
            }

            int row = rowIndex;
            int col = column + 1;
            var reference = r.GetAttribute("r");
            if (reference is not null)
            {
                if (CellReference.TryParse(reference, out var refRow, out var refCol)) (row, col) = (refRow, refCol);
                else Options.Warnings.Warn("xlsx.bad-reference", $"Sheet '{sheet.Name}' has an invalid cell reference '{reference}'.");
            }
            column = col;

            var cell = ReadCell(r, sheet, ref formulaWarned);
            if (cell is null || col >= CellReference.MaxColumns) continue;
            if (++_cellCount > Options.MaxCellCount)
                throw new InvalidDataException($"The workbook has more than the limit of {Options.MaxCellCount} cells.");
            sheet.Cells[(row, col)] = cell;
        }
        r.Read();
    }

    /// <summary>Reads one <c>c</c> element, leaving the reader after its end tag; null for cells without a displayable value.</summary>
    private Cell? ReadCell(XmlReader r, Worksheet sheet, ref bool formulaWarned)
    {
        var type = r.GetAttribute("t") ?? "n";
        var format = FormatOf(r.GetAttribute("s"));
        if (r.IsEmptyElement)
        {
            r.Read();
            return BuildEmpty(format);
        }

        int depth = r.Depth;
        string? value = null;
        RichString? inline = null;
        bool hasFormula = false;
        r.Read();
        while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
        {
            if (r.NodeType != XmlNodeType.Element)
            {
                r.Read();
                continue;
            }
            switch (r.LocalName)
            {
                case "v":
                    if (r.IsEmptyElement) { value = ""; r.Read(); }
                    else value = r.ReadElementContentAsString();
                    break;
                case "f":
                    hasFormula = true;
                    r.Skip();
                    break;
                case "is":
                    inline = r.ReadRichText();
                    break;
                default:
                    r.Skip();
                    break;
            }
        }
        r.Read();
        return Build(type, value, inline, hasFormula, format, sheet, ref formulaWarned);
    }

    /// <summary>An empty cell only matters when it paints something (fill or border).</summary>
    private static Cell? BuildEmpty(CellFormat format) =>
        format.Style.Fill is not null || format.Style.Gradient is not null || format.Style.Borders.HasAny ? new Cell { Style = format.Style } : null;

    private static string CodeOf(CellFormat format) => format.NumberFormatCode ?? NumberFormat.BuiltIn(format.NumberFormatId) ?? "General";

    private static CellStyle WithColor(CellStyle style, Color? color) => color is { } c ? style with { Font = style.Font with { Color = c } } : style;

    private Cell TextCell(RichString rich, CellFormat format)
    {
        var code = CodeOf(format);
        // Runs are kept only when the text is shown as it is; a text format (e.g. "@" with literals) rewrites it.
        if (code == "General") return new Cell { Text = rich.Text, Runs = ResolveRuns(rich, format.Style.Font), Kind = CellValueKind.Text, Style = format.Style };
        var text = rich.Text;
        var shown = NumberFormat.Parse(code, _palette).FormatText(text);
        return new Cell { Text = shown.Text, Kind = CellValueKind.Text, Style = WithColor(format.Style, shown.Color) };
    }

    private Cell TextCell(string text, CellFormat format) => TextCell(new RichString(text, null), format);

    /// <summary>
    /// Resolves each run on top of the cell's font. A run property that is not mentioned is inherited from the cell (Excel's
    /// own files always spell the properties out, so this only matters for hand-made workbooks).
    /// </summary>
    private List<CellTextRun>? ResolveRuns(RichString rich, RunProperties cellFont)
    {
        if (rich.Runs is not { Count: > 0 } runs) return null;
        var result = new List<CellTextRun>(runs.Count);
        foreach (var run in runs)
        {
            if (run.Text.Length == 0) continue;
            var font = run.Properties is { } rPr ? XlsxStyles.ApplyFont(rPr, cellFont, _resolveColor) : cellFont;
            result.Add(new CellTextRun(run.Text, font));
        }
        return result.Count == 0 ? null : result;
    }

    private Cell? Build(string type, string? value, RichString? inline, bool hasFormula, CellFormat format, Worksheet sheet, ref bool formulaWarned)
    {
        var style = format.Style;
        if (type == "inlineStr")
            return inline is null ? BuildEmpty(format) : TextCell(inline, format);
        if (value is null)
        {
            if (hasFormula && !formulaWarned)
            {
                formulaWarned = true;
                Options.Warnings.Warn("xlsx.formula-no-cache", $"Sheet '{sheet.Name}' has formulas without a cached value; those cells are shown empty.");
            }
            return BuildEmpty(format);
        }

        switch (type)
        {
            case "s":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index >= 0 && index < _sharedStrings.Count)
                    return TextCell(_sharedStrings[index], format);
                Options.Warnings.Warn("xlsx.bad-shared-string", $"Sheet '{sheet.Name}' refers to a shared string that does not exist.");
                return BuildEmpty(format);
            case "str":
            case "d":
                return TextCell(XlsxXml.DecodeEscapes(value), format);
            case "b":
                return new Cell { Text = value.Trim() is "1" or "true" ? "TRUE" : "FALSE", Kind = CellValueKind.Bool, Style = style };
            case "e":
                return new Cell { Text = value, Kind = CellValueKind.Error, Style = style };
            default:
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
                {
                    var shown = NumberFormat.Parse(CodeOf(format), _palette).FormatNumber(number, Date1904);
                    return new Cell { Text = shown.Text, Kind = CellValueKind.Number, Style = WithColor(style, shown.Color), Value = number };
                }
                return value.Length == 0 ? BuildEmpty(format) : new Cell { Text = value, Kind = CellValueKind.Text, Style = style };
        }
    }

    // ------------------------------------------------------------------ defined names

    private void ApplyDefinedNames(Worksheet sheet, int sheetPosition, List<DefinedName> names)
    {
        foreach (var name in names)
        {
            if (name.LocalSheetId != sheetPosition) continue;
            if (name.Name == "_xlnm.Print_Area")
            {
                var areas = CellReference.ParseAreas(name.Formula).Where(a => a.Range is not null).ToList();
                if (areas.Count == 0) continue;
                sheet.PageSetup.PrintAreas.Clear();
                sheet.PageSetup.PrintAreas.AddRange(areas.Select(a => a.Range!.Value));
            }
            else if (name.Name == "_xlnm.Print_Titles")
            {
                foreach (var area in CellReference.ParseAreas(name.Formula))
                {
                    if (area.RowSpan is { } rows) sheet.PageSetup.TitleRows = rows;
                    else if (area.ColumnSpan is { } columns) sheet.PageSetup.TitleColumns = columns;
                }
            }
        }
    }
}
